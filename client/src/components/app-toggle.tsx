import { useLocation } from "wouter";
import { cn } from "@/lib/utils";

type AppKey = "ponpower" | "iolm";

interface AppToggleProps {
  active: AppKey;
}

export function AppToggle({ active }: AppToggleProps) {
  const [, navigate] = useLocation();

  const go = (key: AppKey) => {
    if (key === active) return;
    navigate(key === "ponpower" ? "/" : "/iolm");
  };

  return (
    <div className="flex items-center justify-center">
      <div className="inline-flex items-center rounded-full border border-border bg-secondary/40 p-1 shadow-inner">
        <button
          type="button"
          onClick={() => go("ponpower")}
          className={cn(
            "px-5 py-1.5 text-sm font-semibold rounded-full transition-colors",
            active === "ponpower"
              ? "bg-primary text-primary-foreground shadow"
              : "text-muted-foreground hover:text-foreground"
          )}
          aria-pressed={active === "ponpower"}
        >
          PonPower
        </button>
        <button
          type="button"
          onClick={() => go("iolm")}
          className={cn(
            "px-5 py-1.5 text-sm font-semibold rounded-full transition-colors",
            active === "iolm"
              ? "bg-primary text-primary-foreground shadow"
              : "text-muted-foreground hover:text-foreground"
          )}
          aria-pressed={active === "iolm"}
        >
          IOLM
        </button>
      </div>
    </div>
  );
}
