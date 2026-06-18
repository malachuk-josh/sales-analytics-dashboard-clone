import React from "react";

export function Slider({ value = [0], min = 0, max = 100, step = 1, onValueChange, className = "", ...props }) {
  const current = Array.isArray(value) ? value[0] : value;
  return (
    <input
      type="range"
      min={min}
      max={max}
      step={step}
      value={current}
      onChange={(event) => onValueChange?.([Number(event.target.value)])}
      className={`h-2 w-full accent-lime-400 ${className}`}
      {...props}
    />
  );
}
