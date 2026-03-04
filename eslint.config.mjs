export default [
  {
    files: ["**/*.js"],
    languageOptions: {
      ecmaVersion: 2021,
      sourceType: "script"
    },
    rules: {
      "no-unused-vars": "warn",
      "no-undef": "error",
      "no-redeclare": "error"
    }
  }
];
