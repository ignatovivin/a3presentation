import { useEffect, useState, type PropsWithChildren, type ReactNode } from "react";
import { createPortal } from "react-dom";

type StructureDrawerProps = PropsWithChildren<{
  open: boolean;
  onOpenChange: (open: boolean) => void;
  title: ReactNode;
  description?: ReactNode;
  footer?: ReactNode;
}>;

export function StructureDrawer({ open, onOpenChange, title, description, footer, children }: StructureDrawerProps) {
  const [isMounted, setIsMounted] = useState(open);
  const [isVisible, setIsVisible] = useState(open);

  useEffect(() => {
    let timeoutId: number | undefined;

    if (open) {
      setIsMounted(true);
      timeoutId = window.setTimeout(() => setIsVisible(true), 10);
    } else if (isMounted) {
      setIsVisible(false);
      timeoutId = window.setTimeout(() => setIsMounted(false), 260);
    }

    return () => {
      if (timeoutId !== undefined) {
        window.clearTimeout(timeoutId);
      }
    };
  }, [open, isMounted]);

  if (!isMounted) {
    return null;
  }

  return createPortal(
    <div
      className={`drawer-root${isVisible ? " is-open" : ""}`}
      data-testid="structure-drawer"
      role="dialog"
      aria-modal="true"
      aria-label={typeof title === "string" ? title : "Структура"}
    >
      <button type="button" className="drawer-backdrop" aria-label="Закрыть структуру" onClick={() => onOpenChange(false)} />
      <aside className="drawer-panel">
        <div className="drawer-header">
          <div className="drawer-heading">
            <h2 className="drawer-title">{title}</h2>
            {description ? <p className="drawer-description">{description}</p> : null}
          </div>
          <button type="button" className="drawer-close" aria-label="Закрыть" onClick={() => onOpenChange(false)}>
            ×
          </button>
        </div>
        <div className="drawer-body">{children}</div>
        {footer ? <div className="drawer-footer">{footer}</div> : null}
      </aside>
    </div>,
    document.body,
  );
}
