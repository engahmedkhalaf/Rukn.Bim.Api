import { FAQ } from "@/types";

export const faqs: FAQ[] = [
  {
    question: "How do I install a Revit add-in?",
    answer:
      "Download the installer from our Downloads page, run the setup wizard, and restart Revit. The add-in will appear in the Add-Ins tab. For manual installation, copy the .addin file to %AppData%\\Autodesk\\Revit\\Addins\\20XX and the DLL to the specified folder.",
  },
  {
    question: "Which versions of Revit and Navisworks are supported?",
    answer:
      "Our add-ins support Autodesk Revit and Navisworks versions 2022 through 2026. We release compatibility updates within 30 days of each new Autodesk version.",
  },
  {
    question: "How does licensing work?",
    answer:
      "Free tools require no license. Premium tools use a license key tied to your email address. Professional plans include licenses for up to 2 devices. Enterprise plans offer unlimited device licenses.",
  },
  {
    question: "Can I try premium add-ins before purchasing?",
    answer:
      "Yes! We offer a 14-day free trial for all premium add-ins. No credit card required. Contact us to get your trial license key.",
  },
  {
    question: "Do you offer custom BIM development services?",
    answer:
      "Absolutely. We specialize in custom Revit and Navisworks add-in development, BIM automation, and workflow optimization. Contact us to discuss your project requirements.",
  },
  {
    question: "How do I activate my license?",
    answer:
      "Go to the License Activation page, enter your license key and select the product. Click Activate and you will receive a confirmation. You can also activate directly within the add-in.",
  },
  {
    question: "What support options are available?",
    answer:
      "Free users get community forum support. Professional users receive priority email support with 24-hour response time. Enterprise clients get a dedicated account manager and SLA-backed support.",
  },
  {
    question: "Can I transfer my license to another computer?",
    answer:
      "Yes. Deactivate the license on your current computer through the add-in settings, then activate it on the new computer. Professional licenses support up to 2 simultaneous activations.",
  },
];
