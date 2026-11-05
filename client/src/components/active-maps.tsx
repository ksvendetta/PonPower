import { useRef, useState } from "react";
import { Map } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Popover, PopoverContent, PopoverTrigger } from "@/components/ui/popover";
import type { ActiveMap } from "@shared/active-maps";

function formatCreated(iso: string): string {
  const date = new Date(iso);
  if (Number.isNaN(date.getTime())) return "";
  return date.toLocaleString(undefined, {
    month: "short",
    day: "numeric",
    year: "numeric",
    hour: "numeric",
    minute: "2-digit",
  });
}

function launchMap(id: string) {
  const url = `/map.html?id=${encodeURIComponent(id)}`;
  const opened = window.open(url, "_blank", "noopener,noreferrer");
  if (!opened) window.location.assign(url);
}

export function ActiveMapsButton() {
  const [open, setOpen] = useState(false);
  const [maps, setMaps] = useState<ActiveMap[] | null>(null);
  const [error, setError] = useState("");
  const requestRef = useRef(0);

  const load = () => {
    const request = ++requestRef.current;
    setError("");
    setMaps(null);
    fetch("/api/maps")
      .then(async (response) => {
        if (!response.ok) throw new Error("Could not load active maps.");
        return response.json() as Promise<{ maps?: ActiveMap[] }>;
      })
      .then((body) => {
        if (request !== requestRef.current) return;
        setMaps(Array.isArray(body.maps) ? body.maps : []);
      })
      .catch((err: unknown) => {
        if (request !== requestRef.current) return;
        setMaps([]);
        setError(err instanceof Error ? err.message : "Could not load active maps.");
      });
  };

  return (
    <Popover
      open={open}
      onOpenChange={(next) => {
        setOpen(next);
        if (next) load();
      }}
    >
      <PopoverTrigger asChild>
        <Button variant="outline" size="sm" className="gap-2">
          <Map className="w-4 h-4" /> Active Maps
        </Button>
      </PopoverTrigger>
      <PopoverContent align="end" className="w-80 p-0">
        <div className="border-b px-3 py-2">
          <p className="text-sm font-medium">Active terminal maps</p>
          <p className="text-xs text-muted-foreground">Project number and when the map was created.</p>
        </div>
        <div className="max-h-80 overflow-y-auto p-1">
          {maps === null && !error && (
            <p className="px-3 py-4 text-sm text-muted-foreground">Loading maps…</p>
          )}
          {error && (
            <p role="alert" className="px-3 py-4 text-sm text-destructive">{error}</p>
          )}
          {maps && maps.length === 0 && !error && (
            <p className="px-3 py-4 text-sm text-muted-foreground">No active terminal maps.</p>
          )}
          {maps?.map((map) => (
            <button
              key={map.id}
              type="button"
              className="w-full rounded-md px-3 py-2 text-left hover:bg-accent focus-visible:outline-none focus-visible:ring-1 focus-visible:ring-ring"
              onClick={() => {
                setOpen(false);
                launchMap(map.id);
              }}
            >
              <span className="block text-sm font-medium leading-tight break-words">{map.project}</span>
              <span className="mt-0.5 block text-xs text-muted-foreground">{formatCreated(map.createdAt)}</span>
            </button>
          ))}
        </div>
      </PopoverContent>
    </Popover>
  );
}
