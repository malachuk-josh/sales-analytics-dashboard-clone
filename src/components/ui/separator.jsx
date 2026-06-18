import React from "react";

export function Separator({ className = "", ...props }) {
  return <div className={`h-px w-full ${className}`} role="separator" {...props} />;
}
