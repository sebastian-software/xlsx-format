import { Hero, Features } from "ardo/ui"

export default function HomePage() {
  return (
    <>
      <Hero
        name="XLSX Format"
        text="Documentation Made Simple"
        tagline="Focus on your content, not configuration"
        actions={[
          { text: "Get Started", link: "/guide/getting-started", theme: "brand" },
          { text: "GitHub", link: "https://github.com", theme: "alt" },
        ]}
      />
      <Features
        items={[
          {
            title: "Fast",
            icon: "⚡",
            details: "Lightning fast builds with Vite",
          },
          {
            title: "Simple",
            icon: "✨",
            details: "Easy to set up and use",
          },
          {
            title: "Flexible",
            icon: "🎨",
            details: "Fully customizable theme",
          },
        ]}
      />
    </>
  )
}
