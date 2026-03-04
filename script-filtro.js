/**
 * SISTEMA DE GESTÃO DE DESCONTOS E AUXÍLIOS
 * Versão: 10.0 (Diamond Edition - Arquitetura Sênior & Robustez Total)
 */

// =================================================================================
// 1. CONFIGURAÇÃO E CONSTANTES
// =================================================================================

const PLANILHA_DESCONTO = "ITENS EM LOTE";
const PLANILHA_1414 = "ITENS EM LOTE - 1414";
const ABA_INTEGRAL = "DESCONTO INTEGRAL";
const ABA_HISTORICO = "DESCONTOS PUBLICADOS";

const COLUNAS_CFG = {
  LOTE: { CPF: 0, MOTIVO: 5 },
  P1414: { CPF: 0, MOTIVO: 7 }
};

const CORES = {
  INVALIDO_D89:    "#fff2cc", 
  DUPLICATA:       "#f50505", 
  SOBREPOSICAO:    "#cedfff", 
  RESTITUICAO:     "#b6d7a8", 
  ABSORVIDO:       "#b6d7a8", 
  SUSPEITA_FERIAS: "#cccccc", 
  PADRAO:          "#ffffff"
};

const COTA_PARTE_PATENTE = {
  "coronel": 526.50, "tenente-coronel": 517.26, "tenente coronel": 517.26, "major": 509.83,
  "capitão": 420.02, "capitao": 420.02, "1º ten": 379.10, "1 ten": 379.10, "primeiro ten": 379.10, "primeiro-tenente": 379.10,
  "2º ten": 344.39, "2 ten": 344.39, "segundo ten": 344.39, "segundo-tenente": 344.39, "aspirante": 336.34, "aspirante-a-oficial": 336.34,
  "suboficial": 283.67, "1º sgt": 252.12, "1 sgt": 252.12, "primeiro sargento": 252.12, "primeiro-sargento": 252.12,
  "2º sgt": 219.34, "2 sgt": 219.34, "segundo sargento": 219.34, "segundo-sargento": 219.34,
  "3º sgt": 175.87, "3 sgt": 175.87, "terceiro sargento": 175.87, "terceiro-sargento": 175.87,
  "cabo": 115.59, "cb": 115.59, "taifeiro de primeira": 106.92, "t1": 106.92, "taifeiro de segunda": 101.60, "t2": 101.60,
  "soldado de primeira": 88.57, "s1": 88.57, "soldado de segunda": 81.14, "s2": 81.14
};
const KEYS_PATENTES = Object.keys(COTA_PARTE_PATENTE);

const LEGISLACAO_1414_PADRAO = 'na alínea "g" do subitem 6.2 da ICA 161-16/2024 - Programa Auxílio - Transporte';
const LEGISLACAO_1414_HOSPEDAGEM = 'na alínea "d" do subitem 6.2 da ICA 161-16/2024 - Programa Auxílio - Transporte';

const MAPA_LEGISLACAO = {
  "dispensa médica": 'na alínea "c" do subitem 6.2 da ICA 161-16/2024 - Programa Auxílio - Transporte',
  "recompensa": 'na alínea "c" do subitem 6.2 da ICA 161-16/2024 - Programa Auxílio - Transporte',
  "licença paternidade": 'na alínea "a" do subitem 6.2 da ICA 161-16/2024 - Programa Auxílio - Transporte',
  "licença maternidade": 'na alínea "a" do subitem 6.2 da ICA 161-16/2024 - Programa Auxílio - Transporte',
  "tratamento de saúde": 'na alínea "c" do subitem 6.2 da ICA 161-16/2024 - Programa Auxílio - Transporte',
  "detenção": 'na alínea "e" do subitem 6.2 da ICA 161-16/2024 - Programa Auxílio - Transporte',
  "prisão": 'na alínea "e" do subitem 6.2 da ICA 161-16/2024 - Programa Auxílio - Transporte',
  "núpcias": 'na alínea "d" do subitem 6.2 da ICA 161-16/2024 - Programa Auxílio - Transporte',
  "luto": 'na alínea "d" do subitem 6.2 da ICA 161-16/2024 - Programa Auxílio - Transporte',
  "trânsito": 'na alínea "d" do subitem 6.2 da ICA 161-16/2024 - Programa Auxílio - Transporte',
  "instalação": 'na alínea "d" do subitem 6.2 da ICA 161-16/2024 - Programa Auxílio - Transporte',
  "hospedagem": 'na alínea "d" do subitem 6.2 da ICA 161-16/2024 - Programa Auxílio - Transporte',
  "comissionamento": 'na alínea "g" do subitem 6.2 da ICA 161-16/2024 - Programa Auxílio - Transporte',
  "missão": 'na alínea "g" do subitem 6.2 da ICA 161-16/2024 - Programa Auxílio - Transporte',
};

const MESES_NOMES = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"];
const MESES_ABREV = ["JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"];

const FERIADOS_RAW = [
  new Date(2025, 0, 1), new Date(2025, 2, 3), new Date(2025, 2, 4), new Date(2025, 2, 5),
  new Date(2025, 3, 17), new Date(2025, 3, 18), new Date(2025, 3, 21), new Date(2025, 4, 1),
  new Date(2025, 5, 19), new Date(2025, 5, 20), new Date(2025, 7, 15), new Date(2025, 9, 27),
  new Date(2025, 10, 20), new Date(2025, 10, 21), new Date(2025, 11, 17), new Date(2025, 11, 24),
  new Date(2025, 11, 25), new Date(2025, 11, 26),new Date(2025, 11, 27), new Date(2025, 11, 31), 
  new Date(2026, 0, 1), new Date(2026, 1, 16), new Date(2026, 1, 17), //new Date(2026, 1, 18),
  new Date(2026, 3, 3), new Date(2026, 3, 21), new Date(2026, 4, 1), new Date(2026, 5, 4), 
  new Date(2026, 7, 15), new Date(2026, 8, 7), new Date(2026, 9, 12), new Date(2026, 10, 2),  
  new Date(2026, 10, 15), new Date(2026, 10, 20), new Date(2026, 11, 8), new Date(2026, 11, 17), 
  new Date(2026, 11, 24), new Date(2026, 11, 25), new Date(2026, 11, 31) 
];
const FERIADOS_SET = new Set(
  FERIADOS_RAW.map(d => Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd"))
);

// GESTÃO DE CACHE (LRU - Least Recently Used)
var CACHE_NORMALIZACAO = {};
var CACHE_KEYS = [];
var CACHE_DATAS_OBJ = {};
const CACHE_LIMIT = 2000;

function limparCachesExecucao() {
  CACHE_NORMALIZACAO = {};
  CACHE_KEYS = [];
  CACHE_DATAS_OBJ = {};
  console.log("🧹 Caches de execução limpos.");
}

// =================================================================================
// 2. EVENTOS E GATILHOS
// =================================================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ Sistema de Descontos')
    .addItem('🔄 Atualizar Tudo (Forçar)', 'executarProcessamentoGlobalForce')
    .addSeparator()
    .addItem('📁 Arquivar Mês (Fechar Mês)', 'arquivarFechamentoMensal')
    .addToUi();
}

function onEdit(e) {
  var cache = CacheService.getScriptCache();
  if (cache.get("SISTEMA_OCUPADO") === "true") return;

  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  var sheetName = sheet.getName();
  
  if (sheetName !== PLANILHA_DESCONTO && sheetName !== PLANILHA_1414) return;
  
  var row = e.range.getRow();
  if (row < 2) return; 
  
  var col = e.range.getColumn();
  var layout = obterLayout(sheetName);
  var colMotivo = (layout === "1414") ? (COLUNAS_CFG.P1414.MOTIVO + 1) : (COLUNAS_CFG.LOTE.MOTIVO + 1);
  
  // AÇÃO 1: NORMALIZAÇÃO VISUAL
  if (col === colMotivo) {
    var valorAtual = e.range.getValue(); 
    if (valorAtual && typeof valorAtual === 'string') {
       var valorNormalizado = normalizarMotivoLote(valorAtual, layout);
       if (valorNormalizado && valorNormalizado !== valorAtual) {
           e.range.setValue(valorNormalizado);
       }
    }
  }

  // AÇÃO 2: FILA DE PROCESSAMENTO
  if (col === 1 || col === colMotivo) {
    var cpf = sheet.getRange(row, 1).getValue();
    if (cpf) fila_registrarAlteracao(cpf);
  }
}

// =================================================================================
// 3. CONTROLE E FILA
// =================================================================================

function fila_registrarAlteracao(cpf) {
  if (!cpf) return;
  var lock = LockService.getScriptLock();
  try {
    if (lock.tryLock(3000)) {
      const CHAVE = "FILA_CPFS_ALTERADOS";
      var props = PropertiesService.getScriptProperties();
      var filaStr = props.getProperty(CHAVE);
      var fila = filaStr ? JSON.parse(filaStr) : [];
      
      var cpfStr = cpf.toString().trim();
      if (fila.indexOf(cpfStr) === -1) {
        fila.push(cpfStr);
        props.setProperty(CHAVE, JSON.stringify(fila));
      }
    }
  } catch (e) {
    console.error("Erro Lock Fila: " + e.message + "\n" + e.stack);
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function fila_lerELimpar() {
  const CHAVE = "FILA_CPFS_ALTERADOS";
  var props = PropertiesService.getScriptProperties();
  var filaStr = props.getProperty(CHAVE);
  if (!filaStr) return [];
  props.deleteProperty(CHAVE);
  return JSON.parse(filaStr);
}

function trigger_ProcessarFila() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(0)) return; 

  try {
    CacheService.getScriptCache().put("SISTEMA_OCUPADO", "true", 45);
    var cpfsPendentes = fila_lerELimpar();

    if (cpfsPendentes.length === 0) return;
    
    console.log("Processando: " + cpfsPendentes.length + " CPFs.");
    executarProcessamentoGlobal(false, cpfsPendentes); 

  } catch (e) {
    console.error("Erro Trigger: " + e.message + "\nStack: " + e.stack);
  } finally {
    CacheService.getScriptCache().remove("SISTEMA_OCUPADO");
    if (lock.hasLock()) lock.releaseLock();
  }
}

function executarProcessamentoGlobalForce() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    SpreadsheetApp.getUi().alert("Sistema ocupado. Tente novamente.");
    return;
  }
  try {
    CacheService.getScriptCache().put("SISTEMA_OCUPADO", "true", 300);
    limparCachesExecucao(); // [NOVO] Limpeza explícita
    
    executarProcessamentoGlobal(true, null); 
    SpreadsheetApp.getActiveSpreadsheet().toast('Atualização Completa Concluída!');
  } catch (e) {
    SpreadsheetApp.getUi().alert("Erro fatal: " + e.message);
    console.error("Erro Fatal Force: " + e.message + "\nStack: " + e.stack);
  } finally {
    CacheService.getScriptCache().remove("SISTEMA_OCUPADO");
    if (lock.hasLock()) lock.releaseLock();
  }
}

// =================================================================================
// 4. MOTOR DE PROCESSAMENTO (Map-Reduce + Leitura Única)
// =================================================================================

function carregarDadosGlobais(ss) {
  var dadosGlobais = {};
  var abas = [PLANILHA_DESCONTO, PLANILHA_1414];
  
  abas.forEach(nomeAba => {
    var sheet = ss.getSheetByName(nomeAba);
    if (!sheet) return;
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      dadosGlobais[nomeAba] = { valores: [], backgrounds: [], notes: [], range: null, layout: obterLayout(nomeAba) };
      return;
    }
    
    var lastCol = sheet.getLastColumn();
    // Leitura única para evitar ir ao Google Sheets várias vezes
    var range = sheet.getRange(2, 1, lastRow - 1, lastCol);
    
    dadosGlobais[nomeAba] = {
      valores: range.getValues(),
      backgrounds: range.getBackgrounds(),
      notes: range.getNotes(),
      range: range,
      layout: obterLayout(nomeAba)
    };
  });
  return dadosGlobais;
}

function executarProcessamentoGlobal(forcar, listaCpfs) {
  limparCachesExecucao(); 

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var setCpfs = new Set(listaCpfs || []); 

  // 1. LEITURA ÚNICA
  var dadosGlobais = carregarDadosGlobais(ss);
  
  // 2. PRÉ-PROCESSAMENTO (Em memória)
  // O mapeamento de conflitos agora roda PRIMEIRO
  var objConflitosGlobais = mapearConflitosGlobaisMemoria(dadosGlobais);
  
  // Passamos os conflitos para o Integral ignorar o que foi absorvido
  var mapaAbsorvidas = sincronizarDescontoIntegralMemoria(dadosGlobais, ss, objConflitosGlobais); 
  var historicoPublicado = carregarMapaHistorico(ss); 

  var cacheD89 = {}; 
  var cacheDiasUteis = {};

  // 3. PROCESSAMENTO PRINCIPAL
  for (var nomeAba in dadosGlobais) {
    var pacote = dadosGlobais[nomeAba];
    if (pacote.valores.length === 0) continue;

    var absorvidas = mapaAbsorvidas[nomeAba] || {};

    var resultado = processarBatchCompleto(
      pacote.valores, pacote.backgrounds, pacote.notes, 
      absorvidas, ss, nomeAba, 
      cacheD89, cacheDiasUteis, 
      objConflitosGlobais, 
      forcar, setCpfs,
      historicoPublicado
    );

    // 4. ESCRITA DINÂMICA
    if (pacote.range && resultado.valores.length > 0) {
      var sheet = pacote.range.getSheet();
      var numRows = resultado.valores.length;
      var numCols = resultado.valores[0].length; 
      
      var rangeEscrita = sheet.getRange(2, 1, numRows, numCols);
      
      rangeEscrita.setValues(resultado.valores);
      rangeEscrita.setBackgrounds(resultado.cores);
      rangeEscrita.setNotes(resultado.notas);
    }
  }
  
  PropertiesService.getScriptProperties().setProperty('ULTIMA_EXECUCAO', new Date().toLocaleString());
}

function processarBatchCompleto(dados, backgrounds, notes, linhasAbsorvidas, ss, nomeAba, cacheD89, cacheDiasUteis, objConflitosGlobais, forcar, setCpfs, mapaHistorico) {
  if (!dados || dados.length === 0) {
     return { valores: [], cores: [], notas: [] };
  }

  // --- FASE 1: SCANNER & BUILDER ---
  var layout = obterLayout(nomeAba);
  var idxMotivo = (layout === "1414") ? COLUNAS_CFG.P1414.MOTIVO : COLUNAS_CFG.LOTE.MOTIVO; 
  var minCols = (layout === "1414") ? 8 : 7;
  
  var larguraEntrada = dados[0].length;
  var larguraSaida = Math.max(larguraEntrada, minCols); 

  var motivosUnicos = new Set();
  for (var i = 0; i < dados.length; i++) {
    var m = (dados[i][idxMotivo] || "").toString().trim();
    if (m) motivosUnicos.add(m);
  }

  var MAPA_INTELIGENCIA = {};
  
  motivosUnicos.forEach(mRaw => {
    var mNorm = normalizarMotivoLote(mRaw, layout); 
    var datas = extrairDatasComCache(mNorm);        
    
    var info = {
      normalizado: mNorm, datas: null, diasUteis: 0,
      mesIndex: -1, mesNome: "", anoStr: "", legislacao: "", chavePeriodo: ""
    };

    if (datas) {
      info.datas = datas;
      
      var regexMultiplos = /In[íi]cio:\s*(\d{2}\/\d{2}\/\d{4})[\s\S]*?T[ée]rmino:\s*(\d{2}\/\d{2}\/\d{4})/gi;
      var matchMulti;
      var somaDiasPicados = 0;
      var achouMultiplos = false;
      
      while ((matchMulti = regexMultiplos.exec(mNorm)) !== null) {
          achouMultiplos = true;
          var dIniMulti = validarData(matchMulti[1]);
          var dFimMulti = validarData(matchMulti[2]);
          if (dIniMulti && dFimMulti) {
              somaDiasPicados += calcularDiasUteisComRegra(dIniMulti, dFimMulti, mNorm, cacheDiasUteis);
          }
      }
      
      if (achouMultiplos && somaDiasPicados > 0) {
          info.diasUteis = somaDiasPicados;
      } else {
          var matchDiasExp = mRaw.match(/concedido(?:s)?\s*(\d+)\s*dias/i) || mRaw.match(/(\d+)\s*dias\s*de/i);
          if (matchDiasExp) {
              info.diasUteis = parseInt(matchDiasExp[1], 10);
          } else {
              info.diasUteis = calcularDiasUteisComRegra(datas.inicio, datas.fim, mNorm, cacheDiasUteis);
          }
      }

      info.mesIndex = datas.inicio.getMonth();
      info.mesNome = obterNomesMesesIntervalo(datas.inicio, datas.fim);
      info.anoStr = datas.inicio.getFullYear().toString();
      
      var inicioStr = Utilities.formatDate(datas.inicio, Session.getScriptTimeZone(), "dd/MM/yyyy");
      var fimStr = Utilities.formatDate(datas.fim, Session.getScriptTimeZone(), "dd/MM/yyyy");
      info.chavePeriodo = "|" + inicioStr + "|" + fimStr; 
      
      if (layout !== "1414") {
        info.legislacao = encontrarLegislacao(info.normalizado);
      }
    }
    MAPA_INTELIGENCIA[mRaw] = info;
  });

  // --- FASE 2: APLICAÇÃO ---
  var outputValores = [];
  var outputCores = [];
  var outputNotas = [];
  var mapaDuplicatasLocal = {}; 

  for (var i = 0; i < dados.length; i++) {
    var linhaValores = dados[i].slice();
    while (linhaValores.length < larguraSaida) linhaValores.push(""); 
    
    var linhaCores = (backgrounds[i] || []).slice();
    while (linhaCores.length < larguraSaida) linhaCores.push(null);
    
    var linhaNotas = (notes[i] || []).slice();
    while (linhaNotas.length < larguraSaida) linhaNotas.push("");

    var cpf = (linhaValores[0] || "").toString().trim();
    
    if (!forcar && cpf && !setCpfs.has(cpf)) {
      var chaveConf = nomeAba + "|" + i;
      if (objConflitosGlobais.notas[chaveConf]) {
         linhaCores.fill(CORES.SOBREPOSICAO);
         linhaNotas.fill(objConflitosGlobais.notas[chaveConf]);
      }
      outputValores.push(linhaValores);
      outputCores.push(linhaCores);
      outputNotas.push(linhaNotas);
      continue; 
    }

    var motivoRaw = (linhaValores[idxMotivo] || "").toString().trim();
    var infoMotivo = MAPA_INTELIGENCIA[motivoRaw]; 

    if (infoMotivo && infoMotivo.normalizado !== motivoRaw) {
        linhaValores[idxMotivo] = infoMotivo.normalizado;
    }

    var cor = CORES.PADRAO;
    var nota = "";
    var datas = null;
    var diasUteis = 0;
    var chaveHistorico = null;

    if (cpf && infoMotivo && infoMotivo.datas) {
        // Clonamos para não afetar as outras linhas que compartilham o texto
        datas = { inicio: infoMotivo.datas.inicio, fim: infoMotivo.datas.fim };
        diasUteis = infoMotivo.diasUteis;
        
        var chaveConflito = nomeAba + "|" + i;
        var dataFimAjustada = objConflitosGlobais.datasAjustadas[chaveConflito];

        // --- APLICA A NOVA REGRA DE CORTE DE SOBREPOSIÇÃO ---
        if (dataFimAjustada === "ABSORVIDA") {
            diasUteis = 0;
        } else if (dataFimAjustada) {
            diasUteis = calcularDiasUteisComRegra(datas.inicio, dataFimAjustada, infoMotivo.normalizado, cacheDiasUteis);
            datas.fim = dataFimAjustada; 
        }

        var cpfLimpo = cpf.replace(/\D/g, "");
        chaveHistorico = cpfLimpo + infoMotivo.chavePeriodo;

        var d89 = buscarDadosFinanceirosD89Map(cpf, infoMotivo.mesIndex, infoMotivo.anoStr, cacheD89, ss);
        
        if (!d89.encontrado) {
             var mAnt = infoMotivo.mesIndex - 1; var aAnt = parseInt(infoMotivo.anoStr);
             if (mAnt < 0) { mAnt = 11; aAnt--; }
             var mPos = infoMotivo.mesIndex + 1; var aPos = parseInt(infoMotivo.anoStr);
             if (mPos > 11) { mPos = 0; aPos++; }

             var d89Ant = buscarDadosFinanceirosD89Map(cpf, mAnt, aAnt.toString(), cacheD89, ss);
             var d89Pos = buscarDadosFinanceirosD89Map(cpf, mPos, aPos.toString(), cacheD89, ss);

             if (d89Ant.encontrado && d89Pos.encontrado) {
                d89 = d89Pos; 
                cor = CORES.SUSPEITA_FERIAS;
                nota = "ℹ️ Possibilidade de férias.";
             } else {
                cor = CORES.INVALIDO_D89;
                nota = "⚠️ CPF não encontrado na D89.";
             }
        }

        if (layout === "1414") {
          linhaValores[1] = infoMotivo.mesNome;
          linhaValores[2] = infoMotivo.anoStr;
          linhaValores[4] = formatarDiasUteis(diasUteis);
          linhaValores[3] = infoMotivo.normalizado.toLowerCase().includes("hospedagem") ? LEGISLACAO_1414_HOSPEDAGEM : LEGISLACAO_1414_PADRAO;
          
          if (d89.encontrado && diasUteis > 0) {
             var valDia = d89.valor / 22;
             linhaValores[5] = valDia;
             linhaValores[6] = numeroPorExtensoReais(valDia) + " – Valor total: " + (valDia * diasUteis).toLocaleString('pt-BR', {style:'currency', currency:'BRL'});
          } else {
             linhaValores[5] = "D89 N/A"; linhaValores[6] = "Dados não encontrados";
          }
        } else {
          linhaValores[1] = formatarDiasUteis(diasUteis); 
          linhaValores[2] = 1;
          linhaValores[3] = infoMotivo.mesNome; 
          linhaValores[4] = infoMotivo.anoStr;
          if (infoMotivo.legislacao !== "LEGISLAÇÃO NÃO ENCONTRADA") {
              linhaValores[6] = infoMotivo.legislacao;
          }
        }
    }

    var chvConflitoFinal = nomeAba + "|" + i;
    var ehAbsorvidaPorConflito = (objConflitosGlobais.datasAjustadas[chvConflitoFinal] === "ABSORVIDA");
    var corOriginal = backgrounds[i] ? backgrounds[i][0] : null;

    if (mapaHistorico && chaveHistorico && mapaHistorico.has(chaveHistorico)) {
       cor = CORES.DUPLICATA; nota = "🔴 Desconto JÁ PUBLICADO.";
    }
    else if (ehAbsorvidaPorConflito) {
       cor = CORES.ABSORVIDO; 
       nota = objConflitosGlobais.notas[chvConflitoFinal];
    }
    else if (datas && diasUteis === 0) {
       cor = CORES.DUPLICATA; nota = "🔴 ERRO: 0 dias úteis.";
    }
    else if (linhasAbsorvidas && linhasAbsorvidas[i + 2]) {
       cor = CORES.ABSORVIDO; nota = "✅ Absorvido.";
    } 
    // --- LÓGICA DE WORKFLOW MANUAL (A COR VERDE) ---
    else if (corOriginal === CORES.RESTITUICAO) {
       cor = CORES.RESTITUICAO;
       
       // Verifica qual era o problema original que o sistema encontraria
       if (cor === CORES.SUSPEITA_FERIAS || nota === "ℹ️ Possibilidade de férias.") {
           nota = "✅ Constatado férias no mês de referencia.";
       } 
       else if (objConflitosGlobais.notas[chvConflitoFinal]) {
           nota = "✅ Sobreposição verificada.\n(Original: " + objConflitosGlobais.notas[chvConflitoFinal].replace("⚠️ ", "") + ")";
       } 
       else {
           nota = "🟢 Manual.";
       }
    }
    // --- SE NÃO FOI VALIDADO MANUALMENTE, EXIBE OS ALERTAS PADRÕES ---
    else if (nota === "ℹ️ Possibilidade de férias.") {
       // Mantém a cor amarela de invalido/suspeita que foi definida lá em cima na busca da D89
       cor = CORES.SUSPEITA_FERIAS; 
    }
    else if (objConflitosGlobais.notas[chvConflitoFinal]) {
       cor = CORES.SOBREPOSICAO; 
       nota = objConflitosGlobais.notas[chvConflitoFinal];
    }

    if (cpf && datas && chaveHistorico) {
       if (!mapaDuplicatasLocal[chaveHistorico]) mapaDuplicatasLocal[chaveHistorico] = [];
       mapaDuplicatasLocal[chaveHistorico].push(i);
    }

    linhaCores.fill(cor);
    linhaNotas.fill(nota);

    outputValores.push(linhaValores);
    outputCores.push(linhaCores);
    outputNotas.push(linhaNotas);
  }

  for (var k in mapaDuplicatasLocal) {
    if (mapaDuplicatasLocal[k].length > 1) {
      mapaDuplicatasLocal[k].forEach(ix => { 
          outputCores[ix].fill(CORES.DUPLICATA); 
          outputNotas[ix].fill("🔴 Duplicado nesta aba."); 
      });
    }
  }

  return { valores: outputValores, cores: outputCores, notas: outputNotas };
}


function mapearConflitosGlobaisMemoria(dadosGlobais) {
  var timelineGlobal = {}; 
  var mapaConflitos = { notas: {}, datasAjustadas: {} };

  for (var nomeAba in dadosGlobais) {
    var pacote = dadosGlobais[nomeAba];
    var dados = pacote.valores;
    var idxMotivo = (pacote.layout === "1414") ? COLUNAS_CFG.P1414.MOTIVO : COLUNAS_CFG.LOTE.MOTIVO;

    for (var i = 0; i < dados.length; i++) {
      var cpf = (dados[i][0] || "").toString().trim();
      var motivo = (dados[i][idxMotivo] || "").toString().trim();
      if (!cpf || !motivo) continue;

      var motivoNorm = normalizarMotivoLote(motivo, pacote.layout); 
      var datas = extrairDatasComCache(motivoNorm); 

      if (datas) {
        if (!timelineGlobal[cpf]) timelineGlobal[cpf] = [];
        timelineGlobal[cpf].push({
          inicio: datas.inicio.getTime(),
          fim: datas.fim.getTime(),
          motivoResumo: motivo.split('-')[0].substring(0, 30) + "...",
          aba: nomeAba,
          linhaIndex: i 
        });
      }
    }
  }

  for (var cpf in timelineGlobal) {
    var eventos = timelineGlobal[cpf];
    
    // Ordena por início crescente. Se o início for igual, a mais longa vem primeiro (para abraçar a menor)
    eventos.sort(function(a, b) {
        if (a.inicio !== b.inicio) return a.inicio - b.inicio;
        return b.fim - a.fim;
    });

    var eventoDominante = null;

    for (var k = 0; k < eventos.length; k++) {
        var evAtual = eventos[k];
        var chaveAtual = evAtual.aba + "|" + evAtual.linhaIndex;

        // Verifica se a dispensa atual bate de frente com a dispensa "Dominante" anterior
        if (eventoDominante && evAtual.inicio <= eventoDominante.fimEfetivo.getTime()) {
            
            // CÁLCULO 1: ENGOLIMENTO (evAtual está 100% dentro da dominante)
            if (evAtual.fim <= eventoDominante.fimEfetivo.getTime()) {
                mapaConflitos.notas[chaveAtual] = "✅ Totalmente absorvida por: " + eventoDominante.motivoResumo;
                mapaConflitos.datasAjustadas[chaveAtual] = "ABSORVIDA";
            } 
            // CÁLCULO 2: PASSAR O BASTÃO (evAtual começa dentro mas termina depois)
            else {
                var chaveDominante = eventoDominante.aba + "|" + eventoDominante.linhaIndex;
                var dataCorte = new Date(evAtual.inicio);
                dataCorte.setDate(dataCorte.getDate() - 1);
                dataCorte.setHours(0,0,0,0);

                // A dominante perde a força no dia anterior ao início da atual
                eventoDominante.fimEfetivo = dataCorte;
                mapaConflitos.datasAjustadas[chaveDominante] = eventoDominante.fimEfetivo;

                mapaConflitos.notas[chaveAtual] = (mapaConflitos.notas[chaveAtual] || "") + "\n⚠️ Sobreposição com: " + eventoDominante.motivoResumo;
                mapaConflitos.notas[chaveDominante] = (mapaConflitos.notas[chaveDominante] || "") + "\n⚠️ Sobreposição com: " + evAtual.motivoResumo;

                // A atual passa a ser a nova dominante
                eventoDominante = evAtual;
                eventoDominante.fimEfetivo = new Date(evAtual.fim);
            }
        } else {
            // Sem conflito (passou livre)
            eventoDominante = evAtual;
            eventoDominante.fimEfetivo = new Date(evAtual.fim);
        }
    }
  }
  return mapaConflitos;
}

function sincronizarDescontoIntegralMemoria(dadosGlobais, ss, objConflitosGlobais) {
  var linhasParaAbsorver = {};
  linhasParaAbsorver[PLANILHA_DESCONTO] = {};
  linhasParaAbsorver[PLANILHA_1414] = {};
  var consolidado = {};
  var cacheD89 = {}; 
  var cacheDiasLocal = {}; 
  
  for (var nomeAba in dadosGlobais) {
    var pacote = dadosGlobais[nomeAba];
    var dados = pacote.valores;
    var idxMotivo = (pacote.layout === "1414") ? COLUNAS_CFG.P1414.MOTIVO : COLUNAS_CFG.LOTE.MOTIVO;

    for (var i = 0; i < dados.length; i++) {
      var cpf = (dados[i][0] || "").toString().trim();
      var motivo = (dados[i][idxMotivo] || "").toString().trim();
      if (!cpf || !motivo) continue;

      // --- [A MÁGICA DA ABSORÇÃO NO INTEGRAL] ---
      var chaveConflito = nomeAba + "|" + i;
      var dataFimAjustada = objConflitosGlobais ? objConflitosGlobais.datasAjustadas[chaveConflito] : null;

      // Se a dispensa foi 100% engolida, pula ela! Não soma dias e não entra no texto.
      if (dataFimAjustada === "ABSORVIDA") {
          continue; 
      }

      var layout = pacote.layout;
      var motivoNorm = normalizarMotivoLote(motivo, layout); 
      var infoData = extrairDatasComCache(motivoNorm);       
      
      if (!infoData) continue;

      // Se for apenas uma sobreposição parcial (passar o bastão), a data de término também é cortada aqui
      var dataFimReal = (dataFimAjustada instanceof Date) ? dataFimAjustada : infoData.fim;

      var mesIdx = infoData.inicio.getMonth();
      var ano = infoData.inicio.getFullYear().toString();
      var chave = cpf + "||" + mesIdx + "||" + ano;
      
      var diasLinha = 0;
      var regexMultiplos = /In[íi]cio:\s*(\d{2}\/\d{2}\/\d{4})[\s\S]*?T[ée]rmino:\s*(\d{2}\/\d{2}\/\d{4})/gi;
      var matchMulti;
      var achouMultiplos = false;
      
      while ((matchMulti = regexMultiplos.exec(motivoNorm)) !== null) {
          achouMultiplos = true;
          var dIniMulti = validarData(matchMulti[1]);
          var dFimMulti = validarData(matchMulti[2]);
          if (dIniMulti && dFimMulti) {
              diasLinha += calcularDiasUteisComRegra(dIniMulti, dFimMulti, motivoNorm, cacheDiasLocal);
          }
      }

      // Se não era um texto picado, faz a conta com a regra oficial e a data REAL ajustada
      if (!achouMultiplos || diasLinha === 0) {
          diasLinha = calcularDiasUteisComRegra(infoData.inicio, dataFimReal, motivoNorm, cacheDiasLocal);
      }

      // Rastreia as datas mínimas e máximas
      if (!consolidado[chave]) {
        consolidado[chave] = { 
          cpf: cpf, mesIndex: mesIdx, anoStr: ano, dias: 0, 
          motivos: new Set(), origens: [],
          minData: new Date(infoData.inicio.getTime()), 
          maxData: new Date(dataFimReal.getTime()) 
        };
      } else {
        if (infoData.inicio < consolidado[chave].minData) consolidado[chave].minData = new Date(infoData.inicio.getTime());
        if (dataFimReal > consolidado[chave].maxData) consolidado[chave].maxData = new Date(dataFimReal.getTime());
      }
      
      consolidado[chave].dias += diasLinha;
      consolidado[chave].motivos.add(motivoNorm);
      consolidado[chave].origens.push({ sheet: nomeAba, row: i + 2 });
    }
  }

  var dadosFinais = [];
  
  for (var chave in consolidado) {
    var item = consolidado[chave];
    var d89 = buscarDadosFinanceirosD89Map(item.cpf, item.mesIndex, item.anoStr, cacheD89, ss);

    if (!d89.encontrado) {
       var anoInt = parseInt(item.anoStr);
       var mAnt = item.mesIndex - 1; var aAnt = anoInt;
       if (mAnt < 0) { mAnt = 11; aAnt--; }
       var mPos = item.mesIndex + 1; var aPos = anoInt;
       if (mPos > 11) { mPos = 0; aPos++; }

       var d89Ant = buscarDadosFinanceirosD89Map(item.cpf, mAnt, aAnt.toString(), cacheD89, ss);
       var d89Pos = buscarDadosFinanceirosD89Map(item.cpf, mPos, aPos.toString(), cacheD89, ss);
       if (d89Ant.encontrado && d89Pos.encontrado) d89 = d89Pos;
    }

    if (d89.encontrado) {
      var valorDiaria = d89.valor / 22;
      var desconto = valorDiaria * item.dias;
      var cotaParte = 0;
      var postoLower = d89.posto.toLowerCase();
      for (var j = 0; j < KEYS_PATENTES.length; j++) {
         var p = KEYS_PATENTES[j];
         if (postoLower.includes(p)) { cotaParte = COTA_PARTE_PATENTE[p]; break; }
      }

      if ((desconto + cotaParte) > (d89.valor + 1)) {
        var textoCompleto = gerarTextoLegalIntegral(item, d89, cotaParte);
        
        // [NOVO] Formata o mês "janeiro/fevereiro" ou "dezembro/2025, janeiro e fevereiro"
        var stringMeses = formatarMesesParaIntegral(item.minData, item.maxData);
        // [NOVO] O ano final na coluna será sempre o ano de término (ex: 2026)
        var anoFinalStr = item.maxData.getFullYear().toString();

        dadosFinais.push([
          item.cpf, textoCompleto, d89.valor, formatarDiasUteis(item.dias), stringMeses, anoFinalStr
        ]);
        item.origens.forEach(o => linhasParaAbsorver[o.sheet][o.row] = true);
      }
    }
  }

  atualizarAbaDescontoIntegral(dadosFinais, ss);
  return linhasParaAbsorver;
}
function gerarTextoLegalIntegral(item, d89, cotaParte) {
    // Pega o início real e o fim real para montar o período completo
    var mesInicio = item.minData.getMonth() + 1;
    var anoInicio = item.minData.getFullYear();
    var mesFim = item.maxData.getMonth() + 1;
    var anoFim = item.maxData.getFullYear();
    var ultimoDiaMesFim = new Date(anoFim, mesFim, 0).getDate(); 
    
    var strPeriodo = "01/" + mesInicio.toString().padStart(2, '0') + "/" + anoInicio + 
                     " a " + ultimoDiaMesFim.toString().padStart(2, '0') + "/" + mesFim.toString().padStart(2, '0') + "/" + anoFim;
    
    var valIntegralFmt = d89.valor.toLocaleString('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2});
    var valCotaFmt = cotaParte.toLocaleString('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2});

    // --- FILTRO DE ADAPTAÇÃO PARA O DESCONTO INTEGRAL ---
    // Ajusta os textos para combinarem com o seu cabeçalho "...por"
    var listaAdapta = Array.from(item.motivos).map(function(motivo) {
        var mLower = motivo.toLowerCase();
        
        // Se for dispensa ou licença, adiciona "motivo de "
        if (mLower.startsWith("dispensa") || mLower.startsWith("licença")) {
            return "motivo de " + motivo;
        } 
        // Se for gratificação ("Por ter participado..."), remove o "Por "
        else if (mLower.startsWith("por ")) {
            var semPor = motivo.substring(4); // Corta os 4 primeiros caracteres ("Por ")
            // Garante que a próxima letra fique minúscula (ex: "Ter participado" vira "ter participado")
            return semPor.charAt(0).toLowerCase() + semPor.slice(1);
        }
        
        return motivo; // Retorna normal para os outros casos (núpcias, luto, etc)
    });

    var textoMotivo = "";

    if (listaAdapta.length === 1) {
        textoMotivo = listaAdapta[0];
    } else if (listaAdapta.length > 1) {
        var ultimo = listaAdapta.pop();
        textoMotivo = listaAdapta.join(", ") + " e " + ultimo; 
    }

    if (!/[.?!]$/.test(textoMotivo)) textoMotivo += ".";

    return textoMotivo + 
        " Em atendimento à alínea b.3 do subitem 6.2.1 da ICA 161-16/2024 - Programa Auxílio-Transporte, seja realizado o desconto integral do benefício mensal no valor de R$ " + valIntegralFmt + "," +
        " conforme especificado abaixo de seu nome, e seja restituído por meio de receita a anular, o valor do desconto do Auxílio-Transporte, referente a 6% (seis por cento) calculados sobre 22 (vinte e dois) dias do soldo do militar, no valor de R$ " + valCotaFmt + ", realizado no período de " + strPeriodo + ".";
}
function atualizarAbaDescontoIntegral(dados, ss) {
  var sheet = ss.getSheetByName(ABA_INTEGRAL);
  if (!sheet) return;
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 6).clearContent();
  if (dados && dados.length > 0) {
    sheet.getRange(2, 1, dados.length, 6).setValues(dados);
    sheet.getRange(2, 3, dados.length, 1).setNumberFormat("#,##0.00");
  }
}

// =================================================================================
// 5. NORMALIZAÇÃO V10.0 (ENGINE MODULAR ROBUSTA)
// =================================================================================

function normalizarMotivoLote(textoRaw, layout) {
  if (!textoRaw) return "";
  
  // 1. Sanitização
  var textoLimpo = sanitizarTexto(textoRaw);
  var cacheKey = textoLimpo + "|" + (layout || "LOTE");
  
  if (CACHE_NORMALIZACAO[cacheKey]) return CACHE_NORMALIZACAO[cacheKey];
  
  // Verificação rápida se já está padronizado
  if (/\(Ref\.:.*?\)$/i.test(textoLimpo) || /\(Ref\s.*?\)$/i.test(textoLimpo)) {
     atualizarCache(cacheKey, textoLimpo);
     return textoLimpo;
  }

  // 2. Processamento
  var resultado = processarMotivo(textoLimpo, layout);
  atualizarCache(cacheKey, resultado);
  return resultado;
}

function atualizarCache(key, value) {
  if (CACHE_NORMALIZACAO[key]) {
     var index = CACHE_KEYS.indexOf(key);
     if (index > -1) CACHE_KEYS.splice(index, 1);
  }
  CACHE_NORMALIZACAO[key] = value;
  CACHE_KEYS.push(key);
  if (CACHE_KEYS.length > CACHE_LIMIT) {
     var oldestKey = CACHE_KEYS.shift();
     delete CACHE_NORMALIZACAO[oldestKey];
  }
}

function sanitizarTexto(texto) {
  if (!texto) return "";
  var t = texto.toString();
  if (t.normalize) { t = t.normalize("NFC"); } // Unicode Normalization
  t = t.replace(/[\u201C\u201D\u201E\u201F]/g, '"'); 
  t = t.replace(/[\u2018\u2019]/g, "'");             
  t = t.replace(/[\u2013\u2014]/g, "-");             
  t = t.replace(/[\u200B\u00A0\uFEFF]/g, " "); 
  t = t.replace(/\s+/g, " ");
  t = t.trim().replace(/\s,\s/g, ", ").replace(/\s\./g, ".");
  return t;
}

function processarMotivo(texto, layout) {
  var textoNormalizado = corrigirDatasNoTextoInteiro(texto);
  var tipo = detectingTipo(textoNormalizado);
  var dados = extrairDados(textoNormalizado, tipo);
  return montarMotivoPadronizado(tipo, dados, layout, textoNormalizado);
}

function detectingTipo(texto) {
  var tLimpo = removerAcentos(texto).toUpperCase(); 
  
  if (tLimpo.includes("SEGUNDO LISTA DE ATESTADOS") || tLimpo.match(/^(?:MEDICO|ODONTO)/)) return "GSAU_LISTA";
  
  if (tLimpo.includes("NUPCIAS")) return "núpcias";
  if (tLimpo.includes("LUTO")) return "luto";
  if (tLimpo.match(/DETIDO|DETENCAO/)) return "detenção";
  if (tLimpo.match(/PRESO|PRISAO/)) return "prisão";
  
  if (tLimpo.includes("INCAPAZ") || tLimpo.includes("SESSAO")) return tLimpo.includes("TRATAMENTO") ? "licença para tratamento de saúde" : "dispensa médica";
  if (tLimpo.match(/(MEDICA|MEDICO|ATESTADO|SAUDE|ODONTO)/)) return tLimpo.includes("TRATAMENTO") ? "licença para tratamento de saúde" : "dispensa médica";

  if (tLimpo.includes("PATERNIDADE")) return "licença paternidade";
  if (tLimpo.includes("MATERNIDADE")) return "licença maternidade";
  if (tLimpo.includes("HOSPEDAGEM")) return "hospedagem";
  if (tLimpo.includes("TRANSITO")) return "trânsito";
  if (tLimpo.includes("COMISSIONAMENTO")) return "comissionamento";
  if (tLimpo.includes("MISSAO") || tLimpo.includes("VISITA") || tLimpo.includes("TECNICA")) return "missão";
  
  if (texto.match(/dia\(s\)\s*período de/i) || tLimpo.includes("RECOMPENSA")) return "dispensa total do serviço como recompensa";

  return null; 
}

function extrairDados(texto, tipo) {
  var dados = { inicio: null, fim: null, boletim: null, bca: null, complemento: "", diasExplicitos: null };

  var refs = extrairDadosReferencia(texto);
  dados.boletim = refs.boletim;
  dados.bca = refs.bca;

  var matchDiasExp = texto.match(/concedido(?:s)?\s*(\d+)\s*dias/i) || texto.match(/(\d+)\s*dias\s*de/i);
  if (matchDiasExp) {
      dados.diasExplicitos = parseInt(matchDiasExp[1], 10);
  }

  if (tipo === "GSAU_LISTA") {
    var matchGsau = texto.match(/^(?:Médico|Odontológico|Odonto)[\s\.]+(\d{2}\/\d{2}\/\d{4})\s+(\d+)\s+(.*)/i);
    if (matchGsau) {
       dados.inicio = matchGsau[1];
       dados.fim = somarDiasData(matchGsau[1], parseInt(matchGsau[2]) - 1);
       var matchRef = matchGsau[3].match(/per[ií]odo\s+de\s+(\d{2}\/\d{2}\/\d{4})\s+a\s+(\d{2}\/\d{2}\/\d{4})/i);
       if (matchRef) dados.complemento = " no período de " + matchRef[1] + " a " + matchRef[2];
    }
  }
  else if (tipo === "dispensa médica" || tipo === "detenção" || tipo === "prisão" || (tipo && tipo.includes("licença")) || (tipo && tipo.includes("recompensa"))) {
    var matchInicio = texto.match(/(?:Sessão|a contar)\s*(?:nº\s*\d+,?)?\s*(?:de)?\s*(\d{2}\/\d{2}\/\d{4})/i) || texto.match(/(\d{2}\/\d{2}\/\d{4})/);
    var matchDias = texto.match(/(?:por|concedidos|gozo de)\s*(\d+)\s*.*dias/i) || texto.match(/(\d+)\s*\(.*?\)\s*dias/i) || texto.match(/(\d+)\s*dias/i);
    
    if (matchDias && matchInicio) {
       var dias = parseInt(matchDias[1]) || converterExtensoParaNumero(matchDias[1]);
       if (!dias && matchDias[1]) dias = parseInt(matchDias[1].replace(/\D/g,''));
       
       if (dias && matchInicio[1]) {
         dados.inicio = matchInicio[1];
         dados.fim = somarDiasData(dados.inicio, dias - 1);
       }
    }
  }
  
  if (!dados.inicio && (tipo === "núpcias" || tipo === "luto")) {
     var matchOcorrido = texto.match(/(?:ocorrido|casamento|óbito|a contar).*?(?:em|de)\s*(\d{2}\/\d{2}\/\d{4})/i);
     if (matchOcorrido) {
         dados.inicio = matchOcorrido[1];
         dados.fim = somarDiasData(dados.inicio, 7); 
     }
  }

  if (!dados.inicio) {
    var datasSimples = extrairDatas(texto);
    if (datasSimples) { dados.inicio = datasSimples.inicio; dados.fim = datasSimples.fim; }
  }

  if (!dados.complemento) {
     // [CORREÇÃO] Aceita "segundo a lista", "segundo lista", etc.
     var matchExistente = texto.match(/(,?\s*segundo\s*(?:a\s*)?lista de Atestados.*)/i);
     if (matchExistente) {
         var compl = matchExistente[1];
         compl = compl.replace(/\s*\(\s*Ref\.:.*?\)/i, ""); // Remove referências repetidas do final
         dados.complemento = compl; 
     }
  }

  return dados;
}

function montarMotivoPadronizado(tipo, dados, layout, textoOriginal) {
  if (!dados.inicio || !dados.fim) {
     var refSimples = dados.boletim ? " (Ref.: " + dados.boletim + ")" : "";
     return limparTextoBase(textoOriginal) + refSimples;
  }

  if (!tipo) console.warn("⚠️ Motivo não identificado:", textoOriginal);

  var motivoFinal = tipo || "motivo não identificado"; 

  var inicioStr = dados.inicio;
  var fimStr = dados.fim;
  
  if (inicioStr instanceof Date) inicioStr = Utilities.formatDate(inicioStr, Session.getScriptTimeZone(), "dd/MM/yyyy");
  if (fimStr instanceof Date) fimStr = Utilities.formatDate(fimStr, Session.getScriptTimeZone(), "dd/MM/yyyy");

  var textoDatas = " no período de " + inicioStr + " a " + fimStr;

  if (motivoFinal === "GSAU_LISTA") {
    motivoFinal = "dispensa médica";
    if (!dados.complemento) {
        dados.complemento = "segundo a Lista de Atestados Emitidos pelo GSAU-LS";
    } else if (!dados.complemento.toLowerCase().includes("segundo")) {
        var espaco = dados.complemento.startsWith(" ") ? "" : " ";
        dados.complemento = "segundo a Lista de Atestados Emitidos pelo GSAU-LS" + espaco + dados.complemento;
    }
  }
  
  if (motivoFinal === "hospedagem" && textoOriginal.match(/Hotel de Trânsito/i)) {
    motivoFinal = "hospedagem em Hotel de trânsito";
    layout = "LOTE"; 
  }

  var refString = "";
  if (layout === "1414") {
    var principal = dados.boletim ? dados.boletim : (dados.bca ? "Ref. " + dados.bca : "");
    var compRef = (dados.boletim && dados.bca) ? " (Ref.: " + dados.bca + ")" : "";
    if (principal) {
       var prefixo = principal.startsWith("BCA") ? " - Conforme a " : " - Conforme o ";
       refString = prefixo + principal + compRef;
    }
    
    // 1. SE JÁ ESTIVER PADRONIZADO
    if (textoOriginal.match(/^Por ter participado/i)) {
       var base = textoOriginal.replace(/\s*-\s*Conforme[\s\S]*/i, "");
       return base + refString;
    }

    // 2. BLOCO INTELIGENTE PARA INTERVALOS PICADOS
    var matchBloco1414 = textoOriginal.match(/(?:Eventual(?:,\s*|\s+))?(?:por\s+)?(ter\s+[\s\S]*?)(?:,\s*)?no\s+per[ií]odo\s+de\s+([\s\S]*?)(?:(?:,\s*)?sendo\s+considerado)/i);
    
    if (matchBloco1414) {
       var descricao = "Por " + matchBloco1414[1].trim();
       
       var regexLimpaBca = /\s*contida\s+no\s+Boletim\s+(?:Interno\s+)?do\s+Comando\s+da\s+Aeron[áa]utica\s*n[º°o\.]?\s*\d+[\s,]*(?:de[\s,]*)?\d{2}\/\d{2}\/\d{4}/i;
       descricao = descricao.replace(regexLimpaBca, "");
       
       var rawPeriodo = matchBloco1414[2];
       
       var regexDataHora = /(\d{2}\/\d{2}\/\d{4})[\s,\-]*[aà]s?\s*(\d{1,2}(?::\d{2})?)\s*h?/gi;
       var matches;
       var extraidos = [];
       
       while ((matches = regexDataHora.exec(rawPeriodo)) !== null) {
          extraidos.push({ data: matches[1], hora: matches[2] });
       }
       
       var strDatas = "";
       if (extraidos.length > 0 && extraidos.length % 2 === 0) {
           var trechos = [];
           for (var k = 0; k < extraidos.length; k += 2) {
               var dIn = extraidos[k];
               var dOut = extraidos[k+1];
               trechos.push("Início: " + dIn.data + ", às " + dIn.hora + "h; Término: " + dOut.data + ", às " + dOut.hora + "h");
           }
           
           if (trechos.length === 1) {
               strDatas = " - " + trechos[0]; 
           } else {
               var ultimo = trechos.pop();
               strDatas = " - " + trechos.join("; ") + " e " + ultimo; 
           }
       } else {
           strDatas = " - Início e Término: " + rawPeriodo.trim();
       }
       
       return descricao + strDatas + refString;
    }

    // 3. PADRÃO ANTIGO 1414 (Fallback absoluto)
    if (textoOriginal.match(/(?:Início|Inicio|Período|Início e Término)[\s\S]*?(?:término|termino|fim|- Conforme)/i)) {
       var textoDescritivo = limparTextoBase(textoOriginal);
       textoDescritivo = textoDescritivo.replace(/[\s\.,;:\-]*(Início|Inicio)\s*:?\s*/i, " - Início ");
       textoDescritivo = textoDescritivo.replace(/[\s\.,;:\-]*(Término|Termino|Fim)\s*:?\s*/i, ", Término ");
       return textoDescritivo + refString;
    }
  } else {
    var refConteudo = dados.boletim ? dados.boletim : dados.bca;
    if (refConteudo) refString = " (Ref.: " + refConteudo + ")";
  }

  // --- A CORREÇÃO DO ESPAÇO ESTÁ AQUI ---
  // Limpa espaços em branco e vírgulas duplicadas no começo do complemento para grudar perfeitamente
  var complementoFinal = "";
  if (dados.complemento) {
    var extraLimpo = dados.complemento.trim();
    extraLimpo = extraLimpo.replace(/^[\s,]+/, ""); 
    complementoFinal = ", " + extraLimpo; 
  }

  return motivoFinal + textoDatas + complementoFinal + refString;
}

// =================================================================================
// 6. HELPERS
// =================================================================================
function calcularDiasUteisComRegra(inicio, fim, motivoNorm, cacheDias) {
  var dStart = new Date(inicio.getTime());
  dStart.setHours(0,0,0,0);
  var dEnd = new Date(fim.getTime());
  dEnd.setHours(0,0,0,0);

  if (dStart > dEnd) return 0;

  var mLower = (motivoNorm || "").toLowerCase();
  var motivosParaAbater = ["dispensa médica", "licença para tratamento de saúde", "prisão", "detenção", "hotel de trânsito"];
  var aplicarRegra = motivosParaAbater.some(function(termo) { return mLower.includes(termo); });

  if (aplicarRegra) {
      // A REGRA EXATA: O primeiro dia da dispensa é descartado.
      // Avançamos o início da contagem no calendário para o dia seguinte.
      dStart.setDate(dStart.getDate() + 1);
  }

  // Se a dispensa era de apenas 1 dia, ao avançar o dStart ele ultrapassa o dEnd, resultando em 0.
  if (dStart > dEnd) return 0;

  // Agora sim, conta os dias úteis do período restante
  return contarDiasUteisComCache(dStart, dEnd, cacheDias);
}
function extrairDatasComCache(texto) {
  if (!texto) return null;
  if (CACHE_DATAS_OBJ[texto]) return CACHE_DATAS_OBJ[texto];
  
  // Usa o extrairDatas abaixo que serve de Fallback para a lógica V10
  var datas = extrairDatas(texto);
  CACHE_DATAS_OBJ[texto] = datas;
  return datas;
}

function extrairDatas(texto) {
  if (!texto) return null;

  var match1414Novo = texto.match(/no per[ií]odo de\s+(\d{2}\/\d{2}\/\d{4})[\s,]*[aà]s?\s*\d{1,2}(?::\d{2})?\s*h?[\s,]*a\s*(\d{2}\/\d{2}\/\d{4})/i);
  if (match1414Novo) {
     var d1 = validarData(match1414Novo[1]);
     var d2 = validarData(match1414Novo[2]);
     if (d1 && d2) return {inicio: d1, fim: d2};
  }
  
  var match1414 = texto.match(/(?:In[íi]cio)[\s\.,;:\-]*(\d{2}\/\d{2}\/\d{4})[\s\S]*?(?:T[ée]rmino|Fim)[\s\.,;:\-]*(\d{2}\/\d{2}\/\d{4})/i);
  if (match1414) {
     var d1 = validarData(match1414[1]);
     var d2 = validarData(match1414[2]);
     if (d1 && d2) return {inicio: d1, fim: d2};
  }

  // [CORREÇÃO] Aceita espaços múltiplos e não quebra com espaços duplos
  var matchPadrao = texto.match(/no per[ií]odo de\s+(\d{2}\/\d{2}\/\d{4})\s+a\s+(\d{2}\/\d{2}\/\d{4})/i);
  if (matchPadrao) {
    var d1 = validarData(matchPadrao[1]);
    var d2 = validarData(matchPadrao[2]);
    if (d1 && d2) return {inicio: d1, fim: d2};
  }

  // [A MÁGICA DE PROTEÇÃO] Remove as datas soltas do GSAU, Portaria e Boletins antes de caçar as datas do período!
  var textoLimpo = texto.replace(/\(Ref\.:.*?\)/gi, " ")
                        .replace(/- Conforme o Boletim.*?(\d{2}\/\d{2}\/\d{4})/gi, " ")
                        .replace(/Boletim.*?(\d{2}\/\d{2}\/\d{4})/gi, " ")
                        .replace(/Portaria[\s\S]*?(\d{2}\/\d{2}\/\d{4})/gi, " ")
                        .replace(/GSAU-LS[\s\S]*?(\d{2}\/\d{2}\/\d{4})[\s\S]*?(\d{2}\/\d{2}\/\d{4})/gi, " "); 
                        
  var regexTodas = /(\d{2}\/\d{2}\/\d{4})/g;
  var todas = [];
  var m;
  
  while ((m = regexTodas.exec(textoLimpo)) !== null) {
    var d = validarData(m[1]);
    if (d) todas.push(d);
  }
  
  if (todas.length >= 2) {
    todas.sort(function(a, b) { return a.getTime() - b.getTime(); });
    return {inicio: todas[0], fim: todas[todas.length - 1]};
  }
  
  return null;
}

function removerAcentos(s) {
  if (!s) return "";
  return s.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function corrigirDatasNoTextoInteiro(texto) {
  if (!texto) return "";
  var regex = /(\d{1,2})\s*(?:de|\/|-|\s)\s*([a-zA-ZçÇ]+)\s*(?:de|\/|-|\s)\s*(\d{4})/gi;
  return texto.replace(regex, function(matchCompleto, dia, mes, ano) { return padronizarData(dia, mes, ano); });
}

function extrairDadosReferencia(t) {
  var dados = { boletim: null, bca: null };
  var tUpper = t.toUpperCase();
  
  if (!tUpper.includes("BOLETIM") && !tUpper.includes("BOL") && !tUpper.includes("BCA")) {
     return dados;
  }

  var regexBol = /(?:BOLETIM|BOL)\s*(?:INT\.?|INTERNO)?\s*(?:DE\s+)?(OSTENSIVO|PESSOAL|RESERVADO|SIGILOSO|INF\.? PESS\.?|INFORMAÇÕES\s+PESSOAIS|INFORMACOES\s+PESSOAIS)?\s*(?:N[º°o\.]?|NR[\.\s]|N[\.\s])\s*([A-Z0-9\.\-\/]+)(?:[,\s\-]*(?:DE|EM)?\s*(\d{2}\/\d{2}\/\d{4}|[\d]{1,2}\s+de\s+[a-zA-Zç]+\s+de\s+\d{4}))?/i;
  var matchBol = t.match(regexBol);
  if (matchBol) {
    var tipo = "Boletim Interno Ostensivo"; 
    var rawTipo = (matchBol[1] || "").toUpperCase(); 
    if (rawTipo.match(/PESSOAL|INF/)) tipo = "Boletim Interno de Informações Pessoais";
    else if (rawTipo.match(/RESERVADO/)) tipo = "Boletim Interno Reservado";

    var numero = matchBol[2];
    var dataFormatada = "";
    if (matchBol[3]) {
      var matchExtenso = matchBol[3].match(/(\d{1,2})\s+de\s+([a-zA-Zç]+)\s+de\s+(\d{4})/i);
      if (matchExtenso) dataFormatada = ", de " + padronizarData(matchExtenso[1], matchExtenso[2], matchExtenso[3]);
      else {
         var mData = matchBol[3].match(/(\d{2}\/\d{2}\/\d{4})/);
         if (mData) dataFormatada = ", de " + mData[1];
      }
    }
    dados.boletim = tipo + " n° " + numero + dataFormatada;
  }
  
  // [NOVO] Lê BCA longo e aceita erro de digitação "de, [data]"
  var regexBca = /(?:BCA|Boletim\s+(?:Interno\s+)?do\s+Comando\s+da\s+Aeron[áa]utica)\s*(?:N[º°o\.]?)\s*(\d+)(?:[,\s]*(?:de[\s,]*)?(\d{2}\/\d{2}\/\d{4}))?/i;
  var matchBca = t.match(regexBca);
  
  if (matchBca) {
     var dataBca = matchBca[2] ? ", " + matchBca[2] : ""; 
     dados.bca = "BCA n° " + matchBca[1] + dataBca;
  }
  
  return dados;
}

function limparTextoBase(t) {
  if (!t) return "";
  var limpo = t;
  limpo = limpo.replace(/\((?:Ref\.|Boletim|BCA|Conforme)[\s\S]*?\)/gi, " ");
  limpo = limpo.replace(/BOLETIM\s+(?:OSTENSIVO|INTERNO|RESERVADO|PESSOAL|DE\s+INFORMAÇÕES)?.*?DE\s+\d{2}\/\d{2}\/\d{4}/gi, " ");
  limpo = limpo.replace(/Ref\.:\s*BCA[\s\S]*?(?=\s-|\sConforme|$)/gi, " ");
  limpo = limpo.replace(/BCA\s*n[º°o]?\s*\d+[\s\S]*?(?=\s-|\sConforme|$)/gi, " ");
  limpo = limpo.replace(/(?:Ref\.|Boletim|BCA)\s.*$/gi, " ");
  var termosRuido = ["Apresentou", "presentou", "homologada", "homologado", "publicado", "parecer", "conforme", "atestado", "médico", "odontológico", "necessidade de", "onde indica", "Dia\\(s\\):", "Dia:"];
  var regexRuido = new RegExp(termosRuido.join("|"), "gi");
  limpo = limpo.replace(regexRuido, " "); 
  return limpo.replace(/\s{2,}/g, " ").trim();
}

const MAPA_NUMEROS_EXTENSO = {
  "UM": 1, "DOIS": 2, "TRES": 3, "TRÊS": 3, "QUATRO": 4, "CINCO": 5,
  "SEIS": 6, "SETE": 7, "OITO": 8, "NOVE": 9, "DEZ": 10,
  "ONZE": 11, "DOZE": 12, "TREZE": 13, "QUATORZE": 14, "CATORZE": 14, "QUINZE": 15,
  "DEZESSEIS": 16, "DEZESSETE": 17, "DEZOITO": 18, "DEZENOVE": 19,
  "VINTE": 20, "TRINTA": 30
};

function converterExtensoParaNumero(txt) {
  if (!txt) return null;
  var t = removerAcentos(txt).toUpperCase(); 
  for (var chave in MAPA_NUMEROS_EXTENSO) {
    if (t.includes(chave)) return MAPA_NUMEROS_EXTENSO[chave];
  }
  return null;
}

function padronizarData(dia, mes, ano) {
  dia = dia.toString().padStart(2, '0');
  var meses = {'jan': '01', 'fev': '02', 'mar': '03', 'abr': '04', 'mai': '05', 'jun': '06', 'jul': '07', 'ago': '08', 'set': '09', 'out': '10', 'nov': '11', 'dez': '12', 'janeiro': '01', 'fevereiro': '02', 'março': '03', 'abril': '04', 'maio': '05', 'junho': '06', 'julho': '07', 'agosto': '08', 'setembro': '09', 'outubro': '10', 'novembro': '11', 'dezembro': '12'};
  var mesFormatado = mes;
  if (isNaN(mes)) {
    var mesLimpo = mes.toLowerCase().trim().substr(0, 3);
    if (meses[mesLimpo]) mesFormatado = meses[mesLimpo];
  } else { mesFormatado = mes.toString().padStart(2, '0'); }
  return dia + "/" + mesFormatado + "/" + ano;
}
function verificarSeEhDiaUtil(data) {
  var d = new Date(data.getTime());
  var diaSemana = d.getDay();
  // Se for Sábado (6) ou Domingo (0), não é útil
  if (diaSemana === 0 || diaSemana === 6) return false;
  
  // Verifica Feriados
  var dtStr = Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
  if (FERIADOS_SET.has(dtStr)) return false;
  
  return true;
}
function obterNomesMesesIntervalo(dInicio, dFim) {
  var mesesEncontrados = [];
  var atual = new Date(dInicio.getFullYear(), dInicio.getMonth(), 1);
  var fim = new Date(dFim.getFullYear(), dFim.getMonth(), 1);
  var limite = 0;
  while (atual <= fim && limite < 24) {
    mesesEncontrados.push(MESES_NOMES[atual.getMonth()]);
    atual.setMonth(atual.getMonth() + 1);
    limite++;
  }
  return mesesEncontrados.join("/");
}

function contarDiasUteisComCache(inicio, fim, cache) {
  var key = inicio.getTime() + "|" + fim.getTime();
  if (cache && cache[key] !== undefined) return cache[key];
  var count = 0;
  var data = new Date(inicio.getTime());
  data.setHours(0,0,0,0); 
  var fimData = new Date(fim.getTime());
  fimData.setHours(0,0,0,0);
  while (data <= fimData) { 
    var dia = data.getDay();
    if (dia !== 0 && dia !== 6) {
       var dtStr = Utilities.formatDate(data, Session.getScriptTimeZone(), "yyyy-MM-dd");
       if (!FERIADOS_SET.has(dtStr)) count++;
    }
    data.setDate(data.getDate() + 1);
  }
  if (cache) cache[key] = count;
  return count;
}
function formatarMesesParaIntegral(dInicio, dFim) {
  var meses = [];
  var atual = new Date(dInicio.getFullYear(), dInicio.getMonth(), 1);
  var fim = new Date(dFim.getFullYear(), dFim.getMonth(), 1);
  var anoFim = dFim.getFullYear();
  var crossYear = dInicio.getFullYear() !== dFim.getFullYear();

  var limite = 0;
  while (atual <= fim && limite < 24) {
    var nomeMes = MESES_NOMES[atual.getMonth()];
    
    // Se for quebra de ano e o mês pertencer ao ano anterior (ex: 2025), adiciona /2025
    if (crossYear && atual.getFullYear() < anoFim) {
       nomeMes += "/" + atual.getFullYear();
    }
    
    meses.push(nomeMes);
    atual.setMonth(atual.getMonth() + 1);
    limite++;
  }

  // Se tem quebra de ano (ex: dezembro/2025, janeiro e fevereiro)
  if (crossYear) {
    if (meses.length === 1) return meses[0]; // Fallback
    var ultimoMes = meses.pop();
    return meses.join(", ") + " e " + ultimoMes;
  } 
  // Se for mesmo ano (ex: outubro/novembro)
  else {
    return meses.join("/");
  }
}
function validarData(str) {
    if (!str) return null;
    var partes = str.split('/');
    var d = parseInt(partes[0], 10);
    var m = parseInt(partes[1], 10);
    var a = parseInt(partes[2], 10);
    if (m < 1 || m > 12) return null;
    if (d < 1 || d > 31) return null;
    if (a < 1900 || a > 2100) return null;
    return new Date(a, m - 1, d, 12, 0, 0);
}

function somarDiasData(dataStr, dias) {
  if (!dataStr) return null;
  var partes = dataStr.split('/');
  var dia = parseInt(partes[0], 10);
  var mes = parseInt(partes[1], 10) - 1;
  var ano = parseInt(partes[2], 10);
  var data = new Date(ano, mes, dia);
  data.setDate(data.getDate() + dias);
  return data.getDate().toString().padStart(2, '0') + '/' + (data.getMonth() + 1).toString().padStart(2, '0') + '/' + data.getFullYear();
}

function formatarDiasUteis(dias) {
  if (!dias) return "0 dias úteis";
  var extenso = (typeof numeroPorExtensoSimples === 'function') ? numeroPorExtensoSimples(dias) : dias.toString();
  if (typeof extenso === 'string') {
    extenso = extenso.charAt(0).toUpperCase() + extenso.slice(1);
  }
  var sufixo = (dias === 1) ? "dia útil" : "dias úteis";
  return dias + " (" + extenso + ") " + sufixo;
}

function numeroPorExtensoReais(vlr) { 
  if (vlr === 0) return "Zero reais";
  vlr = Math.round(vlr * 100) / 100;
  var partes = vlr.toString().split(".");
  var inteiro = parseInt(partes[0]);
  var centavos = partes[1] ? parseInt(partes[1].padEnd(2, '0').substring(0, 2)) : 0;
  var unidades = ["", "um", "dois", "três", "quatro", "cinco", "seis", "sete", "oito", "nove"];
  var dezenas = ["", "dez", "vinte", "trinta", "quarenta", "cinquenta", "sessenta", "setenta", "oitenta", "noventa"];
  var centenas = ["", "cento", "duzentos", "trezentos", "quatrocentos", "quinhentos", "seiscentos", "setecentos", "oitocentos", "novecentos"];
  function converter(n) {
    if (n < 10) return unidades[n]; 
    if (n < 20) { var especiais = ["dez", "onze", "doze", "treze", "quatorze", "quinze", "dezesseis", "dezessete", "dezoito", "dezenove"]; return especiais[n - 10]; }
    if (n < 100) { var d = Math.floor(n / 10), u = n % 10; return dezenas[d] + (u > 0 ? " e " + unidades[u] : ""); }
    if (n === 100) return "cem";
    if (n < 1000) { var c = Math.floor(n / 100), resto = n % 100; return centenas[c] + (resto > 0 ? " e " + converter(resto) : ""); }
    return n;
  }
  var r = converter(inteiro) + (inteiro === 1 ? " real" : " reais") + (centavos > 0 ? " e " + converter(centavos) + (centavos === 1 ? " centavo" : " centavos") : "");
  return r.charAt(0).toUpperCase() + r.slice(1);
}

function numeroPorExtensoSimples(n) {
  const uni = ["zero","um","dois","três","quatro","cinco","seis","sete","oito","nove","dez","onze","doze","treze","quatorze","quinze","dezesseis","dezessete","dezoito","dezenove"];
  const dez = ["", "", "vinte","trinta","quarenta","cinquenta","sessenta","setenta","oitenta","noventa"];
  if (n < 20) return uni[n];
  if (n < 100) { var d = Math.floor(n/10); var u = n%10; return dez[d] + (u ? " e " + uni[u] : ""); }
  return n.toString(); 
}

function buscarDadosFinanceirosD89Map(cpfRaw, mesIndex, anoStr, cache, ss) {
  if (mesIndex < 0 || mesIndex > 11) return { encontrado: false };
  var cpf = (cpfRaw || "").toString().replace(/\D/g, '');
  var nomeAba = "D89 - " + MESES_ABREV[mesIndex] + "/" + anoStr;
  if (!cache[nomeAba]) {
    cache[nomeAba] = {}; 
    var sheet = ss.getSheetByName(nomeAba);
    if (sheet) {
      var lastRow = sheet.getLastRow();
      if (lastRow >= 2) {
        var dados = sheet.getRange(2, 2, lastRow - 1, 10).getValues();
        for (var i = 0; i < dados.length; i++) {
          var c = (dados[i][0] || "").toString().replace(/\D/g, ''); 
          if (c) {
             var val = (typeof dados[i][9] === 'number') ? dados[i][9] : parseFloat((dados[i][9]||'0').toString().replace(/[^\d,]/g, '').replace(',', '.'));
             cache[nomeAba][c] = { posto: (dados[i][4] || "").toString(), valor: val };
          }
        }
      }
    }
  }
  var item = cache[nomeAba][cpf];
  return item ? { encontrado: true, posto: item.posto, valor: item.valor } : { encontrado: false, valor: 0, posto: "" };
}

function encontrarLegislacao(motivo) {
  var m = motivo.toLowerCase();
  for (var chave in MAPA_LEGISLACAO) if (m.includes(chave)) return MAPA_LEGISLACAO[chave];
  return "LEGISLAÇÃO NÃO ENCONTRADA"; 
}
function carregarMapaHistorico(ss) {
  var mapa = new Set();
  var sheet = ss.getSheetByName(ABA_HISTORICO);
  
  if (!sheet) return mapa;
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return mapa;
  
  // Lê colunas A até F (Indices 0 a 5)
  // Estrutura esperada: [Timestamp, CPF, Motivo, Inicio, Fim, AbaOrigem]
  var dados = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  
  for (var i = 0; i < dados.length; i++) {
    var cpfRaw = (dados[i][1] || "").toString();
    var inicio = dados[i][3]; 
    var fim = dados[i][4]; 
    
    if (cpfRaw && inicio && fim) {
      var cpfLimpo = cpfRaw.replace(/\D/g, "");
      
      // Garante que a data vire string dd/MM/yyyy para a chave funcionar
      var inicioStr = (inicio instanceof Date) ? Utilities.formatDate(inicio, Session.getScriptTimeZone(), "dd/MM/yyyy") : inicio;
      var fimStr = (fim instanceof Date) ? Utilities.formatDate(fim, Session.getScriptTimeZone(), "dd/MM/yyyy") : fim;
      
      // Chave única: CPF + DataInicio + DataFim
      mapa.add(cpfLimpo + "|" + inicioStr + "|" + fimStr);
    }
  }
  return mapa;
}
function arquivarFechamentoMensal() {
  var ui = SpreadsheetApp.getUi();
  var resposta = ui.alert(
    "Fechamento Mensal",
    "Deseja arquivar os dados e LIMPAR as planilhas?\n\nIsso salvará os lançamentos atuais no Histórico e zerará as abas (incluindo cores, notas e a aba de Desconto Integral) para o próximo mês.",
    ui.ButtonSet.YES_NO
  );

  if (resposta !== ui.Button.YES) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Abas que GERAM histórico (apenas a base de dados)
  var abasOrigem = [PLANILHA_1414, PLANILHA_DESCONTO]; 
  var sheetHistorico = ss.getSheetByName(ABA_HISTORICO);

  if (!sheetHistorico) {
    ui.alert("Erro", "A aba de Histórico (" + ABA_HISTORICO + ") não foi encontrada.", ui.ButtonSet.OK);
    return;
  }

  var dadosParaArquivar = [];
  var dataAtual = new Date(); 

  // --- FASE 1: EXTRAIR E SALVAR ---
  abasOrigem.forEach(function(nomeAba) {
    var sheet = ss.getSheetByName(nomeAba);
    if (!sheet) return;

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return; 

    var dados = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    var layout = obterLayout(nomeAba);
    var idxMotivo = (layout === "1414") ? COLUNAS_CFG.P1414.MOTIVO : COLUNAS_CFG.LOTE.MOTIVO;

    for (var i = 0; i < dados.length; i++) {
      var cpf = (dados[i][0] || "").toString().trim();
      var motivoRaw = (dados[i][idxMotivo] || "").toString().trim();

      if (cpf && motivoRaw) {
        var motivoNorm = normalizarMotivoLote(motivoRaw, layout);
        
        var infoData = null;
        if (typeof extrairDatasComCache === "function") {
            infoData = extrairDatasComCache(motivoNorm);
        } else {
            infoData = extrairDatas(motivoNorm);
        }

        if (infoData && infoData.inicio && infoData.fim) {
           var cpfLimpo = cpf.replace(/\D/g, ""); 
           
           dadosParaArquivar.push([
             dataAtual,
             cpfLimpo,
             motivoNorm,
             infoData.inicio,
             infoData.fim,
             nomeAba
           ]);
        }
      }
    }
  });

  if (dadosParaArquivar.length > 0) {
    var primeiraLinhaVazia = sheetHistorico.getLastRow() + 1;
    sheetHistorico.getRange(primeiraLinhaVazia, 1, dadosParaArquivar.length, 6).setValues(dadosParaArquivar);
  }

  // --- FASE 2: FAXINA GERAL ---
  // Incluímos a aba Desconto Integral na lista de limpeza, mas ela não foi arquivada acima
  var abasParaLimpar = [PLANILHA_1414, PLANILHA_DESCONTO, "Desconto Integral"];

  abasParaLimpar.forEach(function(nomeAba) {
    var sheet = ss.getSheetByName(nomeAba);
    if (sheet) {
      var lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        // Pega do início da linha 2 até o fim da aba
        var range = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
        
        range.clearContent();      // Apaga os textos
        range.setBackground(null); // Tira o fundo vermelho/amarelo/verde
        range.clearNote();         // Apaga as notas de "Duplicado", "Absorvido", etc.
      }
    }
  });

  if (dadosParaArquivar.length > 0) {
    ui.alert("Fechamento Concluído!", dadosParaArquivar.length + " lançamentos foram arquivados e as abas foram limpas com sucesso.", ui.ButtonSet.OK);
  } else {
    ui.alert("Aviso", "As planilhas foram limpas, mas nenhum dado válido com datas foi encontrado para salvar no histórico.", ui.ButtonSet.OK);
  }
}
function obterLayout(nomeAba) {
  return (nomeAba === PLANILHA_1414) ? "1414" : "LOTE";
}