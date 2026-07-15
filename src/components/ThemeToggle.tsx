"use client";
import { useState, useEffect } from "react";
import { useTheme } from "next-themes";
import { Sun, Moon, Monitor } from "lucide-react";
import { Button } from "@/components/ui/button";

export function ThemeToggle({ vertical = false }: { vertical?: boolean }) {
  const { theme, setTheme } = useTheme();
  const [mounted, setMounted] = useState(false);

  useEffect(() => {
    setMounted(true);
  }, []);

  if (!mounted) {
    return (
      <div className={`flex items-center gap-0.5 bg-muted/50 border border-border p-0.5 rounded-lg ${vertical ? "flex-col" : ""}`}>
        <div className="h-8 w-8 rounded-md" />
        <div className="h-8 w-8 rounded-md" />
        <div className="h-8 w-8 rounded-md" />
      </div>
    );
  }

  return (
    <div className={`flex items-center gap-0.5 bg-muted/50 border border-border p-0.5 rounded-lg ${vertical ? "flex-col" : ""}`}>
      <Button
        variant={theme === "light" ? "default" : "ghost"}
        size="icon"
        className="h-8 w-8 rounded-md"
        onClick={() => setTheme("light")}
        title="Mode clair"
        aria-label="Mode clair"
      >
        <Sun size={16} />
      </Button>
      <Button
        variant={theme === "dark" ? "default" : "ghost"}
        size="icon"
        className="h-8 w-8 rounded-md"
        onClick={() => setTheme("dark")}
        title="Mode sombre"
        aria-label="Mode sombre"
      >
        <Moon size={16} />
      </Button>
      <Button
        variant={theme === "system" ? "default" : "ghost"}
        size="icon"
        className="h-8 w-8 rounded-md"
        onClick={() => setTheme("system")}
        title="Système"
        aria-label="Système"
      >
        <Monitor size={16} />
      </Button>
    </div>
  );
}