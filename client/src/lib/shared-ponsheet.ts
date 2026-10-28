import { useSyncExternalStore } from "react";

/**
 * The Ponsheet workbook is the same input for PonPower and F2 Exfo, so an
 * upload on either page is published here and adopted by the other one.
 * Module-level state survives wouter navigation between the two pages.
 */
let current: File | null = null;
const listeners = new Set<() => void>();

/** PonPower only reads .xlsx; F2 Exfo also accepts older formats it can't use. */
export function isSharablePonsheet(file: File): boolean {
  return /\.xlsx$/i.test(file.name);
}

export function setSharedPonsheet(file: File | null) {
  if (current === file) return;
  current = file;
  listeners.forEach(listener => listener());
}

export function getSharedPonsheet(): File | null {
  return current;
}

function subscribe(listener: () => void) {
  listeners.add(listener);
  return () => { listeners.delete(listener); };
}

export function useSharedPonsheet(): File | null {
  return useSyncExternalStore(subscribe, getSharedPonsheet, getSharedPonsheet);
}
