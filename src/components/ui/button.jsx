import React from "react";

const variants = {
  default: "bg-slate-100 text-slate-950 hover:bg-white",
  outline: "border border-slate-600 bg-transparent hover:bg-white/10",
  ghost: "bg-transparent hover:bg-white/10",
};

export function Button({ className = "", variant = "default", type = "button", ...props }) {
  return (
    <button
      type={type}
      className={`inline-flex items-center justify-center gap-2 rounded-md px-3 py-2 text-sm font-medium transition disabled:pointer-events-none disabled:opacity-50 ${variants[variant] || variants.default} ${className}`}
      {...props}
    />
  );
}
