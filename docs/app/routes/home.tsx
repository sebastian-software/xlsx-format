import { Hero, Features } from "ardo/ui"
import { Package, Zap, ShieldCheck, Globe, TreePine, ArrowLeftRight, ArrowRight, Github } from "lucide-react"

export default function HomePage() {
  return (
    <>
      <Hero
        name="xlsx-format"
        text="The XLSX library your bundler will thank you for."
        tagline="Zero dependencies. Fully async. TypeScript-first. Works in Node.js and the browser."
        actions={[
          { text: "Get Started", link: "/guide/getting-started", theme: "brand", icon: <ArrowRight size={16} /> },
          { text: "GitHub", link: "https://github.com/sebastian-software/xlsx-format", theme: "alt", icon: <Github size={16} /> },
        ]}
      />
      <Features
        items={[
          {
            title: "Zero Dependencies",
            icon: <Package size={28} strokeWidth={1.5} />,
            details: "No runtime dependencies. No supply chain risk. Nothing to audit, nothing to break.",
          },
          {
            title: "Fully Async",
            icon: <Zap size={28} strokeWidth={1.5} />,
            details: "Streaming ZIP under the hood. Your event loop stays free while Excel files are read and written.",
          },
          {
            title: "TypeScript-First",
            icon: <ShieldCheck size={28} strokeWidth={1.5} />,
            details: "Written in strict TypeScript from day one. Every export is fully typed — no .d.ts bolted on after the fact.",
          },
          {
            title: "Browser-Ready",
            icon: <Globe size={28} strokeWidth={1.5} />,
            details: "Works with Uint8Array and ArrayBuffer. No Node.js APIs needed, no separate browser bundle required.",
          },
          {
            title: "Tree-Shakeable",
            icon: <TreePine size={28} strokeWidth={1.5} />,
            details: "ESM + CJS with named exports. Your bundler drops what you don't use — ship only what you need.",
          },
          {
            title: "Drop-In Migration",
            icon: <ArrowLeftRight size={28} strokeWidth={1.5} />,
            details: "API close to SheetJS by design. Add `await`, switch to named imports, and you're done.",
          },
        ]}
      />
    </>
  )
}
