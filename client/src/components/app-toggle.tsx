import { useLocation } from "wouter";
import { cn } from "@/lib/utils";

type AppKey = "ponpower" | "iolm" | "exfo";

interface AppToggleProps {
  active: AppKey;
}

const PATHS: Record<AppKey, string> = {
  ponpower: "/",
  iolm: "/iolm",
  exfo: "/exfo",
};

const LABELS: Record<AppKey, string> = {
  ponpower: "PonPower",
  iolm: "IOLM",
  exfo: "F2 Exfo",
};

export function AppToggle({ active }: AppToggleProps) {
  const [, navigate] = useLocation();

  const go = (key: AppKey) => {
    if (key === active) return;
    navigate(PATHS[key]);
  };

  return (
    <div className="flex items-center justify-center">
      <div className="inline-flex items-center rounded-full border border-border bg-secondary/40 p-1 shadow-inner">
        {(Object.keys(PATHS) as AppKey[]).map(key => (
          <button
            key={key}
            type="button"
            onClick={() => go(key)}
            className={cn(
              "px-5 py-1.5 text-sm font-semibold rounded-full transition-colors",
              active === key
                ? "bg-primary text-primary-foreground shadow"
                : "text-muted-foreground hover:text-foreground"
            )}
            aria-pressed={active === key}
          >
            {LABELS[key]}
          </button>
        ))}
      </div>
    </div>
  );
}
