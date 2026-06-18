import React from "react";

const SelectContext = React.createContext(null);

function collectItems(children, items = []) {
  React.Children.forEach(children, (child) => {
    if (!React.isValidElement(child)) return;
    if (child.type === SelectItem) {
      items.push({
        value: child.props.value,
        label: child.props.children,
        disabled: child.props.disabled,
      });
      return;
    }
    if (child.props?.children) collectItems(child.props.children, items);
  });
  return items;
}

function findTriggerClassName(children) {
  let result = "";
  React.Children.forEach(children, (child) => {
    if (result || !React.isValidElement(child)) return;
    if (child.type === SelectTrigger) {
      result = child.props.className || "";
      return;
    }
    if (child.props?.children) result = findTriggerClassName(child.props.children);
  });
  return result;
}

export function Select({ value, onValueChange, children }) {
  const items = collectItems(children);
  const triggerClassName = findTriggerClassName(children);

  return (
    <SelectContext.Provider value={{ value }}>
      <select
        value={value}
        onChange={(event) => onValueChange?.(event.target.value)}
        className={`w-full rounded-md border border-slate-600 bg-slate-900/90 px-3 py-2 text-sm text-slate-100 outline-none transition focus:border-lime-400 ${triggerClassName}`}
      >
        {items.map((item) => (
          <option key={item.value} value={item.value} disabled={item.disabled}>
            {item.label}
          </option>
        ))}
      </select>
    </SelectContext.Provider>
  );
}

export function SelectTrigger({ children }) {
  return <>{children}</>;
}

export function SelectValue() {
  const context = React.useContext(SelectContext);
  return <>{context?.value || ""}</>;
}

export function SelectContent({ children }) {
  return <>{children}</>;
}

export function SelectItem({ children }) {
  return <>{children}</>;
}
