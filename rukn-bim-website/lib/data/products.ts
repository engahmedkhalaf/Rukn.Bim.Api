import { Product } from "@/types";

export const products: Product[] = [
  {
    id: "kh-element-navigator",
    slug: "kh-element-navigator",
    name: "KH Element Navigator",
    shortDescription:
      "Advanced element navigation and search tool for Autodesk Navisworks.",
    description:
      "KH Element Navigator is a powerful Navisworks add-in that revolutionizes how you navigate and search elements within your BIM models. With advanced filtering, custom search queries, and intelligent element grouping, this tool dramatically speeds up model review and clash detection workflows.",
    category: "navisworks",
    tier: "premium",
    price: "$149",
    priceNote: "per year",
    version: "2.3.1",
    features: [
      "Advanced multi-criteria element search",
      "Custom search query builder",
      "Intelligent element grouping & classification",
      "One-click element isolation & zoom",
      "Export search results to Excel",
      "Batch element property editing",
      "Saved search presets",
      "Integration with clash detection workflows",
    ],
    systemRequirements: [
      "Autodesk Navisworks 2022–2026",
      "Windows 10/11 (64-bit)",
      ".NET Framework 4.8+",
      "4 GB RAM minimum",
    ],
    changelog: [
      {
        version: "2.3.1",
        date: "2026-06-01",
        changes: [
          "Fixed element grouping performance issue",
          "Added support for Navisworks 2026",
          "Improved search query builder UI",
        ],
      },
      {
        version: "2.2.0",
        date: "2026-03-15",
        changes: [
          "New batch property editing feature",
          "Export results to Excel",
          "Performance improvements",
        ],
      },
    ],
    downloadUrl: "/downloads",
    docsUrl: "/docs",
    icon: "Compass",
  },
  {
    id: "qic-5d-boq-manager",
    slug: "qic-5d-boq-manager",
    name: "QIC 5D BOQ Manager",
    shortDescription:
      "Automated Bill of Quantities generation and 5D BIM cost management for Revit.",
    description:
      "QIC 5D BOQ Manager transforms your Revit models into comprehensive Bills of Quantities with automated cost estimation. Map Revit categories to BOQ codes, generate detailed quantity takeoffs, and export professional reports — all within the Revit environment.",
    category: "revit",
    tier: "premium",
    price: "$199",
    priceNote: "per year",
    version: "1.5.0",
    features: [
      "Automated BOQ generation from Revit models",
      "Custom category-to-BOQ code mapping",
      "5D cost estimation integration",
      "Professional Excel report export",
      "Multi-model quantity consolidation",
      "Audit trail and change tracking",
      "Template-based BOQ formatting",
      "Supabase cloud sync for team collaboration",
    ],
    systemRequirements: [
      "Autodesk Revit 2022–2026",
      "Windows 10/11 (64-bit)",
      ".NET Framework 4.8+",
      "8 GB RAM recommended",
    ],
    changelog: [
      {
        version: "1.5.0",
        date: "2026-05-20",
        changes: [
          "Added cloud sync via Supabase",
          "New audit trail feature",
          "Multi-model consolidation",
        ],
      },
      {
        version: "1.4.0",
        date: "2026-02-10",
        changes: [
          "Template-based BOQ formatting",
          "Improved Excel export",
          "Bug fixes",
        ],
      },
    ],
    downloadUrl: "/downloads",
    docsUrl: "/docs",
    icon: "FileSpreadsheet",
  },
  {
    id: "revit-id-selector",
    slug: "revit-id-selector",
    name: "Selection by Revit ID Tool",
    shortDescription:
      "Quickly select and highlight elements in Revit using Element IDs.",
    description:
      "A lightweight, free Revit add-in that lets you select and highlight elements by their Element ID. Perfect for debugging, clash resolution, and coordination workflows. Paste a list of IDs and instantly navigate to any element in your model.",
    category: "revit",
    tier: "free",
    price: "Free",
    version: "1.2.0",
    features: [
      "Select elements by single or multiple IDs",
      "Paste ID lists from spreadsheets",
      "Auto-zoom to selected elements",
      "Highlight elements with custom colors",
      "Copy element properties to clipboard",
      "Works with linked models",
    ],
    systemRequirements: [
      "Autodesk Revit 2022–2026",
      "Windows 10/11 (64-bit)",
      ".NET Framework 4.8+",
    ],
    changelog: [
      {
        version: "1.2.0",
        date: "2026-04-10",
        changes: [
          "Added linked model support",
          "Improved paste functionality",
          "UI improvements",
        ],
      },
    ],
    downloadUrl: "/downloads",
    docsUrl: "/docs",
    icon: "MousePointerClick",
  },
  {
    id: "bim-utilities-pack",
    slug: "bim-utilities-pack",
    name: "BIM Utilities Pack",
    shortDescription:
      "A collection of essential BIM productivity tools for Revit.",
    description:
      "BIM Utilities Pack bundles essential productivity tools for everyday Revit workflows. Includes batch export, quick parameter editing, view management shortcuts, and family utilities — all in one convenient add-in.",
    category: "revit",
    tier: "free",
    price: "Free",
    version: "3.1.0",
    features: [
      "Batch NWC/IFC/PDF export",
      "Quick parameter editor",
      "View management shortcuts",
      "Family category browser",
      "Sheet index generator",
      "Workset management tools",
    ],
    systemRequirements: [
      "Autodesk Revit 2022–2026",
      "Windows 10/11 (64-bit)",
      ".NET Framework 4.8+",
    ],
    changelog: [
      {
        version: "3.1.0",
        date: "2026-05-01",
        changes: [
          "Added batch PDF export",
          "New workset management tools",
          "Revit 2026 support",
        ],
      },
    ],
    downloadUrl: "/downloads",
    docsUrl: "/docs",
    icon: "Wrench",
  },
];

export function getProductBySlug(slug: string): Product | undefined {
  return products.find((p) => p.slug === slug);
}

export function getProductsByCategory(
  category: "revit" | "navisworks"
): Product[] {
  return products.filter((p) => p.category === category);
}

export function getProductsByTier(tier: "free" | "premium"): Product[] {
  return products.filter((p) => p.tier === tier);
}
