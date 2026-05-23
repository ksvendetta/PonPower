import { Switch, Route, Router as WouterRouter } from "wouter";
import { queryClient } from "./lib/queryClient";
import { QueryClientProvider } from "@tanstack/react-query";
import { Toaster } from "@/components/ui/toaster";
import { TooltipProvider } from "@/components/ui/tooltip";
import NotFound from "@/pages/not-found";
import Home from "@/pages/home";
import HomePub from "@/pages/home-pub";
import ExfoPub from "@/pages/exfo-pub";
import Iolm from "@/pages/iolm";
import Exfo from "@/pages/exfo";
import { PasswordGate } from "@/components/password-gate";

const base = import.meta.env.BASE_URL.replace(/\/$/, "");

function GatedRoutes() {
  return (
    <PasswordGate>
      <Switch>
        <Route path="/">
          <Home />
        </Route>
        <Route path="/iolm" component={Iolm} />
        <Route path="/exfo">
          <Exfo />
        </Route>
        <Route component={NotFound} />
      </Switch>
    </PasswordGate>
  );
}

function App() {
  return (
    <QueryClientProvider client={queryClient}>
      <TooltipProvider>
        <Toaster />
        <WouterRouter base={base}>
          <Switch>
            <Route path="/pub" component={HomePub} />
            <Route path="/pub/exfo" component={ExfoPub} />
            <Route>
              <GatedRoutes />
            </Route>
          </Switch>
        </WouterRouter>
      </TooltipProvider>
    </QueryClientProvider>
  );
}

export default App;
