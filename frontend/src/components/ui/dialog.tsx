import type { PropsWithChildren, ReactNode } from "react";
import { createPortal } from "react-dom";

import { cn } from "@/lib/utils";

type DialogProps = PropsWithChildren<{
  open: boolean;
  onOpenChange: (open: boolean) => void;
}>;

export function Dialog({ open, onOpenChange, children }: DialogProps) {
  if (!open) {
    return null;
  }

  return createPortal(
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4">
      <button
        aria-label="Close dialog"
        className="absolute inset-0 bg-stone-950/50 backdrop-blur-sm"
        onClick={() => onOpenChange(false)}
        type="button"
      />
      <div className="relative z-10 w-full max-w-2xl">{children}</div>
    </div>,
    document.body,
  );
}

export function DialogContent({ className, children }: PropsWithChildren<{ className?: string }>) {
  return <div className={cn("rounded-3xl border bg-card text-card-foreground shadow-2xl", className)}>{children}</div>;
}

export function DialogHeader({ className, children }: PropsWithChildren<{ className?: string }>) {
  return <div className={cn("flex flex-col gap-2 p-6 pb-0", className)}>{children}</div>;
}

export function DialogTitle({ children }: { children: ReactNode }) {
  return <h2 className="text-xl font-semibold tracking-tight">{children}</h2>;
}

export function DialogDescription({ children }: { children: ReactNode }) {
  return <p className="text-sm text-muted-foreground">{children}</p>;
}

export function DialogFooter({ className, children }: PropsWithChildren<{ className?: string }>) {
  return <div className={cn("flex flex-col-reverse gap-2 p-6 pt-0 sm:flex-row sm:justify-end", className)}>{children}</div>;
}
