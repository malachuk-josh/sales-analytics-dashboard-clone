import React from "react";

export const Input = React.forwardRef(function Input({ className = "", ...props }, ref) {
  return (
    <input
      ref={ref}
      className={`w-full rounded-md border border-slate-600 bg-transparent px-3 py-2 text-sm outline-none transition placeholder:text-slate-500 focus:border-lime-400 ${className}`}
      {...props}
    />
  );
});
