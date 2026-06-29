import React from "react";

export function Card({ className = "", ...props }) {
  return <div className={`glass-surface rounded-xl border shadow-sm ${className}`} {...props} />;
}

export function CardHeader({ className = "", ...props }) {
  return <div className={`p-4 ${className}`} {...props} />;
}

export function CardTitle({ className = "", ...props }) {
  return <h3 className={className} {...props} />;
}

export function CardContent({ className = "", ...props }) {
  return <div className={`p-4 ${className}`} {...props} />;
}
