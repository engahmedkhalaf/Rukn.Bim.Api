export interface Product {
  id: string;
  slug: string;
  name: string;
  shortDescription: string;
  description: string;
  category: "revit" | "navisworks";
  tier: "free" | "premium";
  price: string;
  priceNote?: string;
  version: string;
  features: string[];
  systemRequirements: string[];
  changelog: { version: string; date: string; changes: string[] }[];
  downloadUrl: string;
  docsUrl: string;
  icon: string;
}

export interface PricingTier {
  id: string;
  name: string;
  price: string;
  priceNote: string;
  description: string;
  features: string[];
  highlighted: boolean;
  cta: string;
  ctaLink: string;
}

export interface BlogPost {
  id: string;
  slug: string;
  title: string;
  excerpt: string;
  content: string;
  author: string;
  date: string;
  readTime: string;
  category: string;
  tags: string[];
}

export interface FAQ {
  question: string;
  answer: string;
}

export interface Feature {
  icon: string;
  title: string;
  description: string;
}

export interface NavItem {
  label: string;
  href: string;
}

export interface ContactFormData {
  name: string;
  email: string;
  subject: string;
  message: string;
}
