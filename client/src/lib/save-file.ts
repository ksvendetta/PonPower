/**
 * Overwrite the original file if we have a writable handle.
 * Fall back to showSaveFilePicker (Chromium) with the original filename.
 * Fall back again to a simple <a download> with the original filename.
 *
 * Returns "overwritten" | "saved-as" | "downloaded".
 */
export async function saveOrOverwrite(
  contents: string,
  filename: string,
  handle: FileSystemFileHandle | null | undefined,
  mime: string = "application/json"
): Promise<"overwritten" | "saved-as" | "downloaded"> {
  if (handle) {
    const anyHandle = handle as any;
    // Verify / request write permission
    if (typeof anyHandle.queryPermission === "function") {
      let perm = await anyHandle.queryPermission({ mode: "readwrite" });
      if (perm !== "granted" && typeof anyHandle.requestPermission === "function") {
        perm = await anyHandle.requestPermission({ mode: "readwrite" });
      }
      if (perm !== "granted") {
        throw new Error("Permission to overwrite the original file was denied.");
      }
    }
    const writable = await (handle as any).createWritable();
    await writable.write(contents);
    await writable.close();
    return "overwritten";
  }

  const w = window as any;
  if (typeof w.showSaveFilePicker === "function") {
    const newHandle = await w.showSaveFilePicker({
      suggestedName: filename,
      types: [
        {
          description: "JSON",
          accept: { [mime]: [".json"] },
        },
      ],
    });
    const writable = await newHandle.createWritable();
    await writable.write(contents);
    await writable.close();
    return "saved-as";
  }

  // Final fallback
  const blob = new Blob([contents], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  return "downloaded";
}
