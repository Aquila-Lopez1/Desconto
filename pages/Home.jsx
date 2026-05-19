import Navbar from "../components/landing/Navbar";
import HeroSection from "../components/landing/HeroSection";
import HowItWorks from "../components/landing/HowItWorks";
import ScoringSystem from "../components/landing/ScoringSystem";
import WhyJoin from "../components/landing/WhyJoin";
import PricingCards from "../components/landing/PricingCards";
import PriceSimulator from "../components/landing/PriceSimulator";
import PlansComparison from "../components/landing/PlansComparison";
import FAQ from "../components/landing/FAQ";
import WorldCupInfo from "../components/landing/WorldCupInfo";
import Footer from "../components/landing/Footer";

const HERO_BG = "https://media.base44.com/images/public/6a0c82c9d61e67491048f779/24ee91bb9_generated_82aca839.png";
const LOGO = "https://media.base44.com/images/public/6a0c82c9d61e67491048f779/361f698d4_generated_34c58964.png";

export default function Home() {
  return (
    <div className="min-h-screen bg-background">
      <Navbar logoUrl={LOGO} />
      <HeroSection bgUrl={HERO_BG} />
      <HowItWorks />
      <ScoringSystem />
      <WhyJoin />
      <PricingCards />
      <PriceSimulator />
      <PlansComparison />
      <FAQ />
      <WorldCupInfo />
      <Footer logoUrl={LOGO} />
    </div>
  );
}
