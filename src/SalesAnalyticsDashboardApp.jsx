import React, { useEffect, useId, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import pako from "pako";
import { DEFAULT_REPORT_HTML } from "./defaultReportHtml";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Badge } from "@/components/ui/badge";
import { Slider } from "@/components/ui/slider";
import { ScrollArea } from "@/components/ui/scroll-area";
import { Switch } from "@/components/ui/switch";
import { Separator } from "@/components/ui/separator";
import {
  ResponsiveContainer,
  BarChart,
  Bar,
  LineChart,
  XAxis,
  YAxis,
  Tooltip,
  Line,
  ReferenceLine,
  CartesianGrid,
  Cell,
  LabelList,
} from "recharts";
const SimpleIcon = ({ className = "h-4 w-4" }) => (
  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className} aria-hidden="true">
    <circle cx="12" cy="12" r="8" />
    <path d="M12 8v8" />
    <path d="M8 12h8" />
  </svg>
);

const DollarSign = SimpleIcon;
const Target = SimpleIcon;
const TrendingUp = SimpleIcon;
const MessageSquare = SimpleIcon;
const Trophy = SimpleIcon;
const CalendarDays = SimpleIcon;
const Sun = SimpleIcon;
const Moon = SimpleIcon;
const PanelLeftClose = SimpleIcon;
const PanelLeftOpen = SimpleIcon;

const InfoIcon = ({ className = "h-4 w-4" }) => (
  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className} aria-hidden="true">
    <circle cx="12" cy="12" r="9" />
    <path d="M12 11v5" />
    <path d="M12 8h.01" />
  </svg>
);

const DarkModeIcon = ({ className = "h-4 w-4" }) => (
  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className} aria-hidden="true">
    <path d="M20 14.5A7.5 7.5 0 0 1 9.5 4 8.5 8.5 0 1 0 20 14.5Z" />
  </svg>
);

const UploadIcon = ({ className = "h-4 w-4" }) => (
  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className} aria-hidden="true">
    <path d="M12 3v12" />
    <path d="m7 8 5-5 5 5" />
    <path d="M5 15v4h14v-4" />
  </svg>
);

const actionPillStyle = {
  minHeight: "38px",
  borderRadius: "14px",
  border: "1px solid #cbd7e6",
  background: "#f8fbff",
  color: "#111827",
  boxShadow: "0 1px 1px rgba(15, 23, 42, 0.03)",
};

const DEFAULT_DATA_FILE_NAME = "report1781284185572.xls";

const PRELOADED_GROUP_NAMES = ["WMASS", "EMASS", "CT", "Albany", "Exterior"];
const DEFAULT_TEAM_ROSTERS = {
  WMASS: [
    "David Curtis",
    "David Musante",
    "Chris Covington",
    "Emily Nobrega",
    "Stephanie Craven",
    "Caili Stone",
    "Greg Thibeau",
    "Tim Robbins",
    "Colton Duarte",
    "Jon Potter",
    "Heather Lederman",
    "Alex Montanye",
  ],
  EMASS: [
    "Dan O'Keefe",
    "Aaron Newton",
    "Jon Pallone",
    "Liam Cusack",
    "Micah Tully",
    "Matthew Berkovich",
    "Mathias Dos Santos",
    "Jared Berube",
    "Brett Tetreault",
    "Ron Parker",
    "Caleb Rivera",
    "Melissa Tippitt",
    "Collin Conley",
  ],
  CT: [
    "Steve Alkandros",
    "Daniel Lavallee",
    "Michelle Patrick",
    "Ryan Lau",
    "Jessica Maine",
    "Kansas Calloway",
    "Jaydeep Soni",
    "Thomas Martin",
  ],
  Albany: [
    "Jeffrey Scherl",
    "Erika Gibson",
    "Gabriel Stinson",
    "Theresa Fleming",
    "Matthew Kelly",
    "Phil Rogers",
    "Dan DeCesare",
    "Kaelin Loose",
    "Jonathon Rice",
    "Jonathan Rice",
    "Beecher Scarlett",
    "Troy Butler",
    "David Piusienski",
    "Domenic Licciardello",
    "Domonic Licciardello",
  ],
};

const productColors = ["#41cf53", "#7f94c3", "#eeab3a", "#5d9cec", "#9aa6b8"];

const tooltipContentStyle = {
  background: "#2c3542",
  border: "1px solid #4a5568",
  borderRadius: "12px",
  color: "#e6edf7",
};

const tooltipLabelStyle = {
  color: "#e6edf7",
  fontWeight: 600,
};

const tooltipItemStyle = {
  color: "#e6edf7",
};

const KPI_GOALS = {
  nsli: 4000,
  closePct: 0.30,
  demoPct: 0.85,
  netPct: 0.80,
};

const DEFAULT_GOAL_TARGETS = {
  annualNetVolume: 42000000,
  closePct: KPI_GOALS.closePct,
  netPct: KPI_GOALS.netPct,
  nsli: KPI_GOALS.nsli,
  demoPct: KPI_GOALS.demoPct,
  avgTicket: 20000,
};

const DEFAULT_KPI_COLOR_BANDS = {
  closePct: { greenMin: KPI_GOALS.closePct, yellowMin: 0.24 },
  netPct: { greenMin: KPI_GOALS.netPct, yellowMin: 0.70 },
  nsli: { greenMin: KPI_GOALS.nsli, yellowMin: 3500 },
  demoPct: { greenMin: KPI_GOALS.demoPct, yellowMin: 0.74 },
};

const DEFAULT_ANNUAL_VOLUME_BANDS = {
  greenMin: DEFAULT_GOAL_TARGETS.annualNetVolume,
  yellowMin: DEFAULT_GOAL_TARGETS.annualNetVolume * 0.9,
};

const DEFAULT_TIER_THRESHOLDS = { aMin: 0.30, bMin: 0.24 };

const bonusTiers = [
  { min: 125000, max: 149999, rate: 0.01, label: "$125k+" },
  { min: 150000, max: 174999, rate: 0.015, label: "$150k+" },
  { min: 175000, max: Number.POSITIVE_INFINITY, rate: 0.02, label: "$175k+" },
];


function decodeCompressedJson(base64Value) {
  const binary = typeof globalThis !== "undefined" && typeof globalThis.atob === "function"
    ? globalThis.atob(base64Value)
    : typeof Buffer !== "undefined"
      ? Buffer.from(base64Value, "base64").toString("binary")
      : "";
  const bytes = Uint8Array.from(binary, (char) => char.charCodeAt(0));
  return JSON.parse(pako.ungzip(bytes, { to: "string" }));
}

const FALLBACK_DATASETS = {
  "30": { weekLabels: ["1/1/2026 - 1/7/2026"], reps: [] },
  "60": { weekLabels: ["1/1/2026 - 1/7/2026"], reps: [] },
  "90": { weekLabels: ["1/1/2026 - 1/7/2026"], reps: [] },
};

const FALLBACK_DEMO_PCT = {
  "30": {},
  "60": {},
  "90": {},
};

let refreshedDatasets = FALLBACK_DATASETS;
let demoPctByTimeframe = FALLBACK_DEMO_PCT;

try {
  refreshedDatasets = decodeCompressedJson("H4sIAKNT0WkC/+1dXXPbyLH9Kyy95EWhp3um58Nvu/Zmk6y92Vr73jyk9gGW4BXLFOkiKTu+W/7vt2eG+BiAHwA5kuXErnJJgigQOOzpPt19evDHhRQXT/+4+FiW714Ub8r5+uLpvy7kE3iCAvXkzxP5xIRvLy75qG2OgmoOAzXHEZrjiK3jtnXcVcfVk+1pfru8WJXv/Zv/4b+5eHrxvFhM/vGnn8rybcl/tCmLWz74w7835Wq2XPGRRbn53+X87ra8eKqElAhTcXlxNV+uy1+uNhdPxVQpbcPLtj8bDY5/Xs9nF08BrZBTq/jExfzdr8VmtvQv4TOsr1az95tny9v381mxuCq3h2+XH8rbcsEnuvjb7fvlalNe8zV41Oaf+KKBjBNTujSoiPj12gpt+SsISYj8C1BSWf4N3+f71fL67mqz9rB/X2xu1tu3+Odscb38yD85pZXzR54vlyv+uTr5xa/L5dvZ4veLp1ZRfMU/Njfl6uIpCssITOnz58sKvleb8v1NsZiVk2er4kO5aDD858vvXr1KAeSrVih7ACbwOdCmgs8qjVOrc6HnXwOoJSr+xoLU/lLACO3vEYXWZh9wzYXX6IkWcuH7GjXRQoy/T8H6UE6+m78rFter5brB6tnrDlCo2Vp7QGlKLU1QbWmEpKcKc0GFwllG45I06vDVOuNRI/6FR0sa7b/uRss6fnkCFqBQJBJbi+cfgtrz4sPsevLsbrWZrQ+aF1jNH2IPNSNNGzULweAiagZpCpDTwKwA7X0EkrF+aaJRBgJyWul9iLFxEaSL0xpn2oBZCS4BDDU5txe074rVcjH5ufy4WbbW5A87QDNWOPLLvg0aWtRt0EhaqkDTlp2Clbl9mgIif8f+b9j8w71tXdQQbwYAKizsCjEl+NNN3Fl11t2I/b18+3ZVfpq8uuJfzhvMvpu/KRafOqBpCYa6liYFJaA54aACTRqnpsbktDQt2Eb8V16TATSLIToBuP0hgL0ryQ5wVijbAg6tDGt9yNr84XY2/zT5eflmVf5eHF6cGnTf97PZu2RxOlV7fxLsMkw+l7a1Lwla+eswwqG/baWDz5AM315TY28HlKDGkddie31qHR32kDBwtdxsJq/L67fL1XUD2v8s3i2WHxcd2NhRUsfSgqeoMIs/BMBQaxUsPp+NhZAppbXeQyA5DpXkD+6GSTIogUHUMKFSbFvUcmOKUpjYS5r9UL2YFbfs+tfF1bvDXowcKtMPmLJtXFpVSCm/CHQ2WoaGF1FYkTJ8JWVwyy1CDNgLWHPZZ/GKv7Or/6WYz5eLoxyWF4GW/Ilgx9+DoDZUqsXC2AfT1MmsrgsM8OrBYF9OOl46B8xK7AXIMKtIQarPXCEFTvmz9/jESzaqxaY87LNIsgvv+Sxsx8bW+lNIvPzyBUbjnNquQQKyMvAuXoPeaZHzlrbHWzGjTb2VRRn+bLRhFavyevJ9ubp7c9yyOPFwPQePiVkZzk9qtADN1BPafGZFbEneKSswMvyMTsGh9ScGLj0tjVAtlKyynuG2kHpZbPhXHyc/lfP5pyPkAZ112LMq/pdEQjI1eSDGagrZyAMzpUBRiZOhQBg4EG5hulQm3NkgxlXfx0ireu7zxfnkRfGBfVZZ7k+D2Ktq01t/TKQxMSqGqqZZmnl0vjQo0CqKBJ4/A+mjoOTMLBwHbxM0CCnOCZIw6ABSsIyDELl34vXsZjVbT54tP/BLEzq/y2UJ60LM6TBTTIzLoayNK0Ro5XJBZjAsQea70tVsXqmYNEqz12k5YQIdaxNTIdu8dHvmIRb2jA1rNnm1SQLiTrSU0v2ChCCVGFhcmltSyq7V5C1JOAcqQmb8tSCqYB38Re/NGJ3b5oc1XFseO3YxvuZj5bqY/GVe3vqXHnRdTgmtjnkudlyqSXs023VWuBSnO8HCODoHuISqLG03VNoYvd/Ro5ZicLazmr0rJj/O3qzbq3AXUJb5k+qxLKlNihSFhHFL3o2Z2mzkHZwQUK8/qQNinLWCv6j9fHT7Vy2rshHdmsArlPHGBuD1cnZV3Exe3yUhcQd7t8K4nlmhA7PPrNBJM3XZzErCdsWhcrYGzVp9yKzqiz6Luv9ULNYFO3d2WMXH4tP+WMhUV/VqWxhDX+Op6iKNImCE8mU42pkKmW3wC3VBJOsNZg8PrS75LIR+LN6sPF14tZktjq47Iy32LWk3ROiYrjudO11Gzzq9n7IkxSEDqq/1LHh+uZnNJ78ufy9X6yPQINh+LsMXm6TIusn7JHIy77KWE9hqpA4hT6gKL0kk9par6os+Lz8u12t2RZOXxWxxgG6SClfSXWKQ2A/TpcYNEZMBlTWFYQ/k89dLNs6QTB20oPqCz0KHU5ebGfug58v15BXnxO3GxO5Uj/Osfp2dU0xoA8XRpg5unMraqTV5aQDZqlwMII3LkedVJ62LB/HELbB+/VQsOHO5O2BFnLTIoxypRkYHzpezAaFCPwuUIhtahRLGZSsc8Em0kpUxmcpyzpx78vyuWB0prShOBbCf2Qm3GyXDaR3azKzbunZE8xnL3v5MdbVnlunWN5Mf57Ord+vbYnGk8Cu1BjWs7lu/NB861SnFMRe0f2XVVzWMKN6U83k5+aXYrGbtWm+vMwoiVvzadqOTIhPnBXWDD1R8dd4Q37CgcDF7esbVlZ5lM69ntxzd37xh7nO0U2CBDlKfdnD3RQKlstZ0UXiG6LQw44xm+xcVMtvzDCLP5Wxzwzaz+r9PVzdH4DFoDi2nlrPhwGXy28z2/U9cTNXVDwDlr8u7dXkEDM7q7DDfUr0yLxjbs44Ao7qO0QXsT9dl+X7yarmYHcGk51US66gBuUd3cjZ9ObRSVquZb23fzTflfyEQv32+vNA7FGjYKNCwUaBho0DDlgINWwo0bCnQsKVAw0SB9lXJ2wyioh5zJXBpC6cOrdL47kE28YziXFTHalXIgIhibw+cjIWWBxa/AWCqflPO/1FX/TZM0KU0Sqn6/WmX1iaEagRdPppTtk6G5gTA8P2RdUGIpSO4zmgbRREudO5P0nv5hMelWTpnxo70iYKvUZpCpayBvijTiUSK45WEFbJGc9KM2ZJRQgoVUg0EsQkpRBXcpI7Z12mqw+bWzvL9g6VgkgST9p4UzMnEAWh0dV7PyauZ2qztNkneEj1kbJOhViRF6GafrBND1SXiyMle5MoVvdJEY7RigyWJ/Klb2SuV8Gra2+tlMgxTk13DY7eogjOuOniSVBGEEfZcraJRof53ovpOcma1Q30Hqdag9qTKS2nz1TEBRYhPShgTRT8mxiltVLBTTrPoVHkexzSXlsx5EVjdXvaKoINuFmGQFISxO9TRb6SaY0WqEYIC36jN1jkGq1ybBQJ/zraqgT6kYkhbx9ZKJ8kckUM69MUdJu2/W7S6ybA4/ZU6n3Kbgvkpwx9oWNShi3zpRxECvTpDBsmuWFCq7HYkgw+prZPJ2uAu/QhdA6LRfUV8TzRjbbPsFS8Tly02SYXh/SXIQGzIq33r3P40zYO3eJsud0kKThQ9DNe1MZhkRF/XJtJJDAXY9PGlrx9l7XRwEKCor3HRf2oRm7BNq3GU7I0t3GJapeWoB+3IZF2gY7lFuoxnTye4V6Qr2DIyak9jdKdQnlZWU6dqO0K9a42vvLXVu4q9vz5DvTtGPsjJJ8apkVS4lA5XtRUTPgBPXbamAHr2622OKXm3IHFvwkLXERayhyXPDs5qzzGUwvZLwZykJTyeI0ANpePsJG+r16i4vAPRZD4Ysh6kQIHO6N1Vp6jjDRMmmUIKMbWquSajgadLnjiaE/XJZmxutpIiaqmAYQr5dDwkKx2P5qS21vQ4Gxj1aDUUGEcqFdppI+SJiqihQjtGERQdDd+oG82rz6Yxb8TRNlAi4GgSxCtCBvE9Ox6rTxXioeCsyeVQ4g1XlyGbuBdPd+VlSqcrXDVppdd+TSmbWZJwwfzQR/EAJYWMjCORCLSHQJ4gP5PQlcR09WeSnfRw/dlwUTYvI6P68hgl7b6JLg65egr5FFYczppsXUauaWJBNpNgu7nHkZY5eArHOaP7IJq0fiSaRS4texpnMkuMQoBAZhNV8nj6fE59Qw8nYQPrRF+ohSZV1yRCLZ92gMs6SSFilohM/EPWyCzSRddpaKzQrbmjs1AcPjagOebbXh1TpqJug6ZRAjLx1dnMkD8PjESb0+lYd4tMJxaGThoo4HzSdpkQu4d2nBk5VDBYqsOBO+ayB0cKNFBLnqvtVFK+yc0gzbn0JMLGnFrWSxtkXMpj5DxWW5dSoOrUdZjZvkVWBSYoX9LqxxiTpji2mTaQVbDLLFFlAlSl2AoiwRwtzWzu5jz14UCNOHDw2iESFxp3ijiY28kp5hsedrGH0P5/kmq8uY2zUDsuQ3RW96fxpAa1R2DHrsTkVbIasiE8xDI5+H1SqnxwvEQR5bZ/2GQt2q/wk2SKw0VVToPsuT6E1OhsS1POMTEzoVEqan/AxD0iwIZW7OVIzZVv1VGnC0ZJi3aE8Gq4Xs+auDgO7PjSpM6IQBkpdezAMterFqyIm6qMkvHVN/AwYx6WfDeZehOP6ZiHMbYhf7y88o7RJk6Oc2HYkpkzBj/Y7R1YrGcr922gXL3R43Tjl+2swzY8sFWo7HHV8VptG9YYMX99Dw+nMdYwXGMsrMioikTqx9MzdMfV6U7SHm+LBN+Xq3fLD7OjQlJ2o10F8j7QgCNDRplc9c5H/u/Zm6O67LPM60W55JTh/dwvy2Ox0zg3cOOS6pWZFbfxrKNQShGqrmu02PLrkSV3/3+TKR8A5+Fky163sdxsytVYrS59Aa3uNxHzFwbm+fJ2uZhdTV7Mrq5mxeqaqfnyGz55pgA6qrL/DGS8HN7tkMPDVkg++fMEnoCotOb8PbSOm9Zx2xxH1RxHao7LllD+m9q+2fNAAam+2j6dZfN6wEZw71WPOmtxgUwUL3qNbN1ReuxCfLLglYd0iiBX87tp7BckMN0ep6ll8735kcKMey/5EpHv5UEoWHvNWNDhccoSe3+NbiKfaleiFNjZk9XyH7R9Bm3VQ2Nku4MHIEhr7Ncjlaa0Y03NNqOESk4V5Z0ZR4VxV1uo8nRt9L0NRiCzP+pkqozxwwxGsJnJfu1SuVTPa7VqmTp6uX/Gto1sl9vIxu3C7m9gornlhxmYUN7P9cvDrtOqbYl9tfQtv6yqFml8U5stlKxol07yD1JoiO/UsmbjCNwZgxRjZP8KgaQas28ee1I5RcgoraaoAnRBlep7EjKW+VQ02nucDADF4VZ3CvgOtHuA0QCtdo8GJIMBdUWaPYrO2HoLLqLTfnssowLIC6AzKjBG1C55RakdTbpUMGOtbPZi9fEEM25c61xT2ZdBNGrZqQW51j0I3qUXP9uDgndwypjM6kOmpVL1G+/SpTOCrW3+tN+cVGcjH8wvg8RSaaFiFDMhunkxawPgPagT+RTaqBzqxOHTBZ5r0tGdppoGlpQi45adYIysQ6CM9Uilok3dz9iBZaJqunvRk27PHaATaqhRD5818k4Pes1CSVZ2nnfQbArHDEFn26B421FQVoX5Ana0Vaf1XoeQyCqtDw4hEXPylIVkmlHgsLNLCUXp3k62mTVWhl8PdB+VIr2dU3gccwsOUNHpAzRG2G5Wvrep5rxfxKxyR/5g68atJ5mq5hk5BmuAXyFkOlljdNAInjpZM26EgYMu7JidsyYV8Lmm1cRWk1HazBlR5VrZTALSfuHHqJZ5vgF90QQ7I7QgT93ydbDoFEmjkceUG1o2A4rofV6+Z0NpzkWhLT4gCnZ8H2pUQ92NB0GfpUYdqjBHvljXQ1liulWsVk3GLTXfVj52ISCmdNKGVDhKTanWLGRVnzc3e1Y5Y8RcCduK3DVXkk5BaNk8SSOUFSnjDLipK3PO1RB+sXETdtxhH2dR/VP5Bf8ocde+oyp9mFD70QmSzD1IAdhxyka5pEKpP+scQHOjD7xf6Z6BSHJ79y1XnI1l5hg9pcXjHZD8muamemLORzxHNVTrzu5GY/+ZbLGO3HomWzOfz/ZiMz4rK44/+JzO1IoXlFHMlk0K39zlWe5gxCCLC/tU9cmD2jtZBRyNpMzXEdFRDypcEPnauG3Wlm7lnXFhoq2FOjjkws44JC6Z9yF1FJPODhHWZo9mXoInnPlAlhCio2LSH0IWB9LYnTayqk9k0tSDH4LqEAgnkpI8s+zBJfnhg0TGousndJC29nSLo0HmoQSpnd0ppc8yWVTf3j3PyPin8lB/s55ONuGT/0Y3zrdtZO7gxUZCLoUx4+yM13N0SQJo404bnhk++wGEpv+gKyXSsnsiy/cPjtX5OtDC1GUbxikIINjbA9Td5jOHQ5o7PJfRjpKgAwH1d8FOx1vRNM8sAp/Zu3zNjEaGrqn5/ix5enNLD/GkBwA2hh1zXTZxoJzjN0VH9hM6XyeZsymzB6KzhkiaG/vvHFd6ZONL36/KUAPfrMribr45NgtgfZLdA1Omzx+pkVT8QWdLo6r3HrCAjy9mTvNdKhQhqaX9AuNMkQXt5PaAkmnqvVniI5pwGjOP4itr8uBAShtDbclkVqTG6io6d8T4hgyqVOc5e1hl8DwGaNs3QbBgd09d+ydc22wmWL15ljXcIYzVff0Hz4v1LOxLzY9lGI3Chx6N+mKjUn8tC3948qK8LlfHY8SjnHt5sDmh0Mzk1O3Z3dc6IvRtpOpx4vXVz+Z9m0V7JEj99vnz/wMA6TMHLowAAA==");
} catch (error) {
  console.error("Embedded dataset decode failed", error);
}

try {
  demoPctByTimeframe = decodeCompressedJson("H4sIAKNT0WkC/6WWXU/bMBSG/4qVm91UKI4Tf3A3ysY0Cpsof8BtD8QiTZCTFHWI/75ju4WtdrSL3aHj4/h8PO9LXzOWZ+ev2Wdtu5bcwsvQtdl5fiZlyWfZvLamJ/NuZ9rH44EQePAdHh4s7MlyXYNtfJyJAi90DeaRy1HbAVxYYfBStwYastA73TTgw5JW5Sz7sjXNntx2KwuP2odLJn0++fHpGuDB5wpXyXdtYUMuwI6rEBQCM2/0UBvdk8uuJ0vdDl3vzzgXeAPr+IkPdm14scoxOMeAIcvhEFR5Ubr3dmZDbsYevxByOaVYnTVPmlyZVX/oPGdsll3plXXdLAfTHg9cefsNwDNZdq3JzulZ7kbU92atyY024THBqZpl17rtsWJXiH7R+9C1rLAXTK7J/dg0ISgK4YM14NCwk8Ga9ZM74b69n7VpyF33CDb07NPv9ji6hR59mhDY23KAHZDPzZNuNzaMB5/Diu/NFq+vVtiFnwQTWMPC6C2Z4yDCSzLn/gvPtdsgmVu9g9AxL3BA97h76DX52sAWAQnf9ksZangh1/DeSnUc8ny0gwlVCK7ckvqaXDXYWb/V7WFy37qxh8Pfy3U3DOQeNg+d3Rxi12CGGidif+3X9XvMWuNwHJuww/xtlnGPdhI/kcSPKZaEXkheJuAOWog64wH6CCqp4rELJmLgmQPtlHgZI82lmNSAokUaVlXJtH7dhQhaWThZ+4dxEWCnxk09lKf8MEk/eMBWnnCo7ys7BV755BTwrFAR8LKg5d/AC1ZWE7gqydK44ijSzpfSDXXbjlyhLGWKeem2fWI3UnE6wfxlt0XrWJOFWa+Nthv8ThfIU1VaBQvoEMXnxm2sP3pObEMJvxHKWeGfKosVhepRXj1JNVRFAvsyeGkkNrRqmtKD4GxizmVqLazCoi8s+EEMFjSid1hAGRt7QUVKsH4rse654mmxBJOKxZILkRJLWaR5d/5aFBHyKmdFGnmRs1O6C/cPL0W3ymkVGwsX4m/doo68EZ/qKK+mrNyZYYQ7zWms87LC+X0DdCKwZAEbsB9kp9g9NTJZJUxQyQmtRL8FFJ32QUHfqcQRz8ejsPJ/KWBakzJ3mpywwAkn+D+1vr39Bmc3vKWoCQAA");
} catch (error) {
  console.error("Embedded demo pct decode failed", error);
}

const demoWeeks = refreshedDatasets["90"]?.weekLabels || FALLBACK_DATASETS["90"].weekLabels;
const demoReps = refreshedDatasets["90"]?.reps || FALLBACK_DATASETS["90"].reps;

function getDatasetDateBoundsFromWeeks(weekLabels = []) {
  const parsed = (weekLabels || [])
    .map((label) => parseWeekLabelRange(label))
    .filter((item) => item.start instanceof Date && !Number.isNaN(item.start.getTime()) && item.end instanceof Date && !Number.isNaN(item.end.getTime()));

  if (!parsed.length) {
    return { minDate: "", maxDate: "" };
  }

  const minDate = parsed.reduce((min, item) => (item.start < min ? item.start : min), parsed[0].start);
  const maxDate = parsed.reduce((max, item) => (item.end > max ? item.end : max), parsed[0].end);

  return {
    minDate: toIsoDate(minDate),
    maxDate: toIsoDate(maxDate),
  };
}

const embeddedDemoBounds = getDatasetDateBoundsFromWeeks(refreshedDatasets["90"].weekLabels);

function parseMoney(value) {
  const text = cleanText(value).split("$").join("").split(",").join("");
  const n = Number(text);
  return Number.isFinite(n) ? n : 0;
}

function parsePercent(value) {
  const text = cleanText(value).split("%").join("");
  const n = Number(text);
  return Number.isFinite(n) ? n / 100 : 0;
}

function parseUsDate(value) {
  const text = cleanText(value);
  if (!text) return null;
  const parts = text.split("/").map((item) => Number(item));
  if (parts.length !== 3 || parts.some((item) => !Number.isFinite(item))) return null;
  return new Date(parts[2], parts[0] - 1, parts[1]);
}

function toIsoDate(date) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return "";
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function getMostRecentSunday() {
  const d = new Date();
  d.setHours(0, 0, 0, 0);
  d.setDate(d.getDate() - d.getDay());
  return toIsoDate(d);
}

const DEFAULT_DATE_START = "2026-01-01";
const DEFAULT_DATE_END = getMostRecentSunday();

function formatDisplayDate(value) {
  if (!value) return "";
  const date = typeof value === "string" && value.includes("-") ? new Date(`${value}T00:00:00`) : parseUsDate(value);
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return cleanText(value);
  return `${date.getMonth() + 1}/${date.getDate()}/${date.getFullYear()}`;
}

function clampIsoDate(iso, minIso = "", maxIso = "") {
  if (!iso) return "";
  const value = new Date(`${iso}T00:00:00`);
  if (Number.isNaN(value.getTime())) return "";
  const min = minIso ? new Date(`${minIso}T00:00:00`) : null;
  const max = maxIso ? new Date(`${maxIso}T00:00:00`) : null;
  if (min && value < min) return minIso;
  if (max && value > max) return maxIso;
  return iso;
}

function buildLookbackRange(days, anchorEndIso, minIso = "", maxIso = "") {
  const safeEnd = clampIsoDate(anchorEndIso || maxIso || toIsoDate(new Date()), minIso, maxIso || toIsoDate(new Date()));
  const endDate = safeEnd ? new Date(`${safeEnd}T00:00:00`) : new Date();
  const startDate = new Date(endDate);
  startDate.setDate(startDate.getDate() - (Math.max(1, safeNum(days)) - 1));

  if (minIso) {
    const minDate = new Date(`${minIso}T00:00:00`);
    if (!Number.isNaN(minDate.getTime()) && startDate < minDate) {
      return { start: minIso, end: safeEnd || toIsoDate(endDate) };
    }
  }

  return {
    start: toIsoDate(startDate),
    end: safeEnd || toIsoDate(endDate),
  };
}

function startOfWeek(date) {
  const d = new Date(date);
  const day = d.getDay();
  d.setDate(d.getDate() - day);
  d.setHours(0, 0, 0, 0);
  return d;
}

function endOfWeek(date) {
  const d = startOfWeek(date);
  d.setDate(d.getDate() + 6);
  return d;
}

function formatWeekRangeFromDate(date) {
  const start = startOfWeek(date);
  const end = endOfWeek(date);
  return `${start.getMonth() + 1}/${start.getDate()}/${start.getFullYear()} - ${end.getMonth() + 1}/${end.getDate()}/${end.getFullYear()}`;
}

function normalizeProductCategory(product) {
  const text = cleanText(product).toLowerCase();
  if (text.includes("bath")) return "Baths";
  if (text.includes("window")) return "Windows";
  if (text.includes("door")) return "Doors";
  if (text.includes("roof")) return "Roofing";
  return "Siding";
}

function productSelectionParts(selection) {
  const text = cleanText(selection);
  if (!text || text === "All Products") return [];
  return text.split("+").map((part) => cleanText(part)).filter(Boolean);
}

function productSelectionIncludes(selection, product) {
  const normalizedProduct = cleanText(product);
  if (!selection || selection === "All Products") return true;
  const parts = productSelectionParts(selection);
  return parts.includes(normalizedProduct);
}

function productSelectionVolume(products = {}, selection = "All Products") {
  if (!selection || selection === "All Products") {
    return Object.values(products || {}).reduce((sum, value) => sum + safeNum(value), 0);
  }
  return productSelectionParts(selection).reduce((sum, product) => sum + safeNum(products?.[product]), 0);
}

function normalizeProductTotals(products = {}) {
  return {
    Baths: safeNum(products.Baths),
    Windows: safeNum(products.Windows),
    Doors: safeNum(products.Doors),
    Roofing: safeNum(products.Roofing),
    Siding: safeNum(products.Siding) + safeNum(products.Other),
  };
}

function findFallbackRepMeta(repName) {
  const key = canonicalName(repName);
  for (const dataset of Object.values(refreshedDatasets)) {
    for (const rep of dataset.reps || []) {
      if (canonicalName(rep.rep) === key) return rep;
    }
  }
  return null;
}

function extractLabeledValue(text, label) {
  const clean = cleanText(text);
  const index = clean.toLowerCase().indexOf(label.toLowerCase());
  if (index === -1) return "";
  const after = clean.slice(index + label.length).trim();
  const endIndex = after.indexOf("(");
  return cleanText(endIndex >= 0 ? after.slice(0, endIndex) : after);
}

function parseStandardHtmlReport(htmlText) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(htmlText, "text/html");
  const rows = Array.from(doc.querySelectorAll("tr"));
  const headerRow = rows.find((row) => row.querySelectorAll("th").length && row.textContent.includes("Net Amt Split") && row.textContent.includes("Demo%"));
  if (!headerRow) return null;

  const headers = Array.from(headerRow.querySelectorAll("th")).map((cell) => cleanText(cell.textContent));
  const headerIndex = (label) => headers.findIndex((item) => cleanText(item).toLowerCase() === label.toLowerCase());
  const issueIndex = headerIndex("Issue Split");
  const demoCountIndex = headerIndex("Demo Split");
  const soldIndex = headerIndex("Sold Split");
  const soldPriceIndex = headerIndex("Sold Price Split") >= 0 ? headerIndex("Sold Price Split") : headerIndex("SOLD price Split");
  const netCountIndex = headerIndex("Net Split");
  const netAmtIndex = headerIndex("Net Amt Split");
  const workingAmtIndex = headerIndex("Working Amt Split") >= 0 ? headerIndex("Working Amt Split") : headerIndex("Working amt Split");
  const closeIndex = headerIndex("Close%");
  const netIndex = headerIndex("Net%");
  const nsliIndex = headerIndex("Updated NSLI");
  const demoIndex = headerIndex("Demo%");
  const comIndex = headerIndex("COM") >= 0 ? headerIndex("COM") : Math.max(0, headers.length - 1);

  let currentDate = "";
  let currentProduct = "Other";
  let currentTeam = "Unknown";
  const entries = [];
  const dateSet = new Set();

  rows.forEach((row) => {
    const text = cleanText(row.textContent);
    if (!text) return;

    if (row.className.includes("breakRowClass0") && text.includes("Appointment Date:")) {
      currentDate = extractLabeledValue(text, "Appointment Date:");
      if (currentDate) dateSet.add(currentDate);
      return;
    }

    if (text.includes("Sales Team:")) {
      currentTeam = cleanTeamName(extractLabeledValue(text, "Sales Team:"));
      return;
    }

    if (row.className.includes("breakRowClass1") && text.includes("Product Category:")) {
      currentProduct = normalizeProductCategory(extractLabeledValue(text, "Product Category:"));
      return;
    }

    if (row.className.includes("breakRowClass2") && text.includes("Sales Rep:")) {
      const repName = cleanRepName(extractLabeledValue(text, "Sales Rep:"));
      const metricsRow = row.nextElementSibling;
      if (!repName || !metricsRow || !validRepName(repName)) return;
      const cells = Array.from(metricsRow.querySelectorAll(":scope > td")).map((cell) => cleanText(cell.textContent));
      const metricAt = (index) => index >= 0 ? cells[index + 1] : "";
      const fallback = findFallbackRepMeta(repName) || {};
      entries.push({
        date: currentDate,
        rep: repName,
        team: currentTeam && currentTeam !== "Unknown" ? currentTeam : cleanTeamName(fallback.team),
        product: currentProduct,
        issueCount: parseMoney(metricAt(issueIndex)),
        demoCount: parseMoney(metricAt(demoCountIndex)),
        soldCount: parseMoney(metricAt(soldIndex)),
        soldPriceAmount: parseMoney(metricAt(soldPriceIndex)),
        netCount: parseMoney(metricAt(netCountIndex)),
        netVolume: parseMoney(metricAt(netAmtIndex)),
        workingAmount: parseMoney(metricAt(workingAmtIndex)),
        closePct: parsePercent(metricAt(closeIndex)),
        netPct: parsePercent(metricAt(netIndex)),
        nsli: parseMoney(metricAt(nsliIndex)),
        demoPct: parsePercent(metricAt(demoIndex)),
        comPct: parsePercent(metricAt(comIndex)),
        marketingScore: 1 - parsePercent(metricAt(comIndex)),
        talkRatio: safeNum(fallback.talkRatio),
        scriptCompliance: safeNum(fallback.scriptCompliance),
        movement: fallback.movement || "Imported",
      });
    }
  });

  const sortedDates = Array.from(dateSet).sort((a, b) => parseUsDate(a) - parseUsDate(b));
  return {
    sheetNames: ["Standard HTML Report"],
    rawEntries: entries,
    availableDates: sortedDates,
    minDate: sortedDates.length ? toIsoDate(parseUsDate(sortedDates[0])) : "",
    maxDate: sortedDates.length ? toIsoDate(parseUsDate(sortedDates[sortedDates.length - 1])) : "",
  };
}

function groupImportedEntries(
  rawEntries,
  startIso,
  endIso,
  selectedGroup = "All Groups",
  selectedGroupMembers = new Set(),
  selectedRep = "All Reps",
  selectedProduct = "All Products"
) {
  const startDate = startIso ? new Date(`${startIso}T00:00:00`) : null;
  const endDate = endIso ? new Date(`${endIso}T00:00:00`) : null;
  const grouped = {};

  (rawEntries || []).forEach((entry) => {
    const date = parseUsDate(entry.date);
    if (!(date instanceof Date) || Number.isNaN(date.getTime())) return;
    if (startDate && date < startDate) return;
    if (endDate && date > endDate) return;
    if (selectedGroup !== "All Groups" && !selectedGroupMembers.has(canonicalName(entry.rep))) return;
        if (selectedRep !== "All Reps" && canonicalName(entry.rep) !== canonicalName(selectedRep)) return;

    const product = normalizeProductCategory(entry.product);
    if (!productSelectionIncludes(selectedProduct, product)) return;

    const repKey = canonicalName(entry.rep);
    const dateKey = toIsoDate(date);
    const groupKey = selectedProduct === "All Products"
      ? `${dateKey}|${repKey}`
      : `${dateKey}|${repKey}|${product}`;

    if (!grouped[groupKey]) {
      grouped[groupKey] = {
        dateKey,
        weekKey: formatWeekRangeFromDate(date),
        rep: cleanRepName(entry.rep),
        team: cleanTeamName(entry.team),
        netVolume: 0,
        workingAmount: 0,
        soldPriceAmount: 0,
        issueCount: 0,
        demoCount: 0,
        soldCount: 0,
        netCount: 0,
      };
    }

    grouped[groupKey].netVolume += safeNum(entry.netVolume);
    grouped[groupKey].workingAmount += safeNum(entry.workingAmount);
    grouped[groupKey].soldPriceAmount += safeNum(entry.soldPriceAmount);
    grouped[groupKey].issueCount = Math.max(grouped[groupKey].issueCount, safeNum(entry.issueCount));
    grouped[groupKey].demoCount = Math.max(grouped[groupKey].demoCount, safeNum(entry.demoCount));
    grouped[groupKey].soldCount = Math.max(grouped[groupKey].soldCount, safeNum(entry.soldCount));
    grouped[groupKey].netCount = Math.max(grouped[groupKey].netCount, safeNum(entry.netCount));
  });

  return Object.values(grouped);
}

function aggregateRawEntriesToDataset(rawEntries, startIso, endIso) {
  const startDate = startIso ? new Date(`${startIso}T00:00:00`) : null;
  const endDate = endIso ? new Date(`${endIso}T00:00:00`) : null;
  const filteredEntries = (rawEntries || []).filter((entry) => {
    const date = parseUsDate(entry.date);
    if (!(date instanceof Date) || Number.isNaN(date.getTime())) return false;
    if (startDate && date < startDate) return false;
    if (endDate && date > endDate) return false;
    return true;
  });

  const weekStarts = Array.from(new Set(filteredEntries.map((entry) => toIsoDate(startOfWeek(parseUsDate(entry.date)))))).sort();
  const weekLabels = weekStarts.map((iso) => formatWeekRangeFromDate(new Date(`${iso}T00:00:00`)));
  const weekIndexMap = Object.fromEntries(weekStarts.map((iso, index) => [iso, index]));
  const repMap = {};

  filteredEntries.forEach((entry) => {
    const key = canonicalName(entry.rep);
    if (!key) return;
    if (!repMap[key]) {
      repMap[key] = {
        rep: cleanRepName(entry.rep),
        team: cleanTeamName(entry.team),
        movement: entry.movement || "Imported",
        weekly: Array(weekLabels.length).fill(0),
        products: { Baths: 0, Windows: 0, Doors: 0, Roofing: 0, Siding: 0 },
        issueCount: 0,
        demoCount: 0,
        soldCount: 0,
        netCount: 0,
        _talkSamples: [],
        _complianceSamples: [],
        _marketingSamples: [],
      };
    }

    const weekIso = toIsoDate(startOfWeek(parseUsDate(entry.date)));
    const index = weekIndexMap[weekIso];
    if (index >= 0) repMap[key].weekly[index] += safeNum(entry.netVolume);
    repMap[key].products[normalizeProductCategory(entry.product)] += safeNum(entry.netVolume);
    repMap[key].issueCount += safeNum(entry.issueCount);
    repMap[key].demoCount += safeNum(entry.demoCount);
    repMap[key].soldCount += safeNum(entry.soldCount);
    repMap[key].netCount += safeNum(entry.netCount);
    pushIfValue(repMap[key]._talkSamples, entry.talkRatio);
    pushIfValue(repMap[key]._complianceSamples, entry.scriptCompliance);
    pushIfValue(repMap[key]._marketingSamples, entry.marketingScore);
  });

  const reps = Object.values(repMap).map((rep) => {
    const netVolume = rep.weekly.reduce((sum, value) => sum + safeNum(value), 0);
    const closePct = rep.issueCount > 0 ? rep.soldCount / rep.issueCount : 0;
    const netPct = rep.soldCount > 0 ? rep.netCount / rep.soldCount : 0;
    const nsli = rep.issueCount > 0 ? netVolume / rep.issueCount : 0;
    const talkRatio = average(rep._talkSamples, 0);
    const scriptCompliance = average(rep._complianceSamples, 0);
    const demoPct = rep.issueCount > 0 ? rep.demoCount / rep.issueCount : 0;
    const marketingScore = average(rep._marketingSamples, 0);
    return {
      rep: rep.rep,
      team: rep.team,
      netVolume,
      closePct,
      netPct,
      nsli,
      talkRatio,
      scriptCompliance,
      demoPct,
      marketingScore,
      issueCount: rep.issueCount,
      demoCount: rep.demoCount,
      soldCount: rep.soldCount,
      netCount: rep.netCount,
      avgNetPerNetSale: rep.netCount > 0 ? netVolume / rep.netCount : 0,
      movement: rep.movement === "Imported" ? movementFromMetrics({ netVolume, closePct, nsli }) : rep.movement,
      weekly: rep.weekly,
      products: rep.products,
    };
  }).sort((a, b) => b.netVolume - a.netVolume);

  return { reps, weekLabels, minDate: startIso, maxDate: endIso };
}

function getTimeframeWeeks(timeframe) {
  if (timeframe === "30") return refreshedDatasets["30"].weekLabels.length;
  if (timeframe === "60") return refreshedDatasets["60"].weekLabels.length;
  return refreshedDatasets["90"].weekLabels.length;
}

function getRepDemoPct(repName, timeframe) {
  const map = demoPctByTimeframe[timeframe] || {};
  return safeNum(map[cleanRepName(repName)]);
}

function closestEmbeddedTimeframe(days) {
  const value = Math.max(1, safeNum(days));
  if (value <= 45) return "30";
  if (value <= 75) return "60";
  return "90";
}

function getEmbeddedRepBaseline(repName, days, timeframeSets) {
  const timeframeKey = closestEmbeddedTimeframe(days);
  const sourceDataset = timeframeSets?.[timeframeKey] || refreshedDatasets[timeframeKey] || refreshedDatasets["90"];
  const sourceRep = (sourceDataset?.reps || []).find((rep) => canonicalName(rep.rep) === canonicalName(repName));

  if (!sourceRep) {
    return {
      closePct: 0,
      netPct: 0,
      nsli: 0,
      demoPct: getRepDemoPct(repName, timeframeKey),
    };
  }

  return {
    closePct: safeNum(sourceRep.closePct),
    netPct: safeNum(sourceRep.netPct),
    nsli: safeNum(sourceRep.nsli),
    demoPct: safeNum(sourceRep.demoPct) || getRepDemoPct(repName, timeframeKey),
  };
}

function safeNum(value) {
  const n = Number(value);
  return Number.isFinite(n) ? n : 0;
}

function currency(value) {
  return new Intl.NumberFormat("en-US", {
    style: "currency",
    currency: "USD",
    maximumFractionDigits: 0,
  }).format(safeNum(value));
}

function pct(value, digits = 1) {
  return `${(safeNum(value) * 100).toFixed(digits)}%`;
}

function cleanText(value) {
  return String(value || "").split(String.fromCharCode(160)).join(" ").trim();
}

function cleanRepName(value) {
  let text = cleanText(value);
  if (text.startsWith("Sales Rep:")) text = cleanText(text.slice(10));
  if (text.startsWith("Sales Rep 1: Staff Name:")) text = cleanText(text.slice(24));
  const paren = text.indexOf("(");
  if (paren >= 0) text = cleanText(text.slice(0, paren));
  return text;
}

function cleanTeamName(value) {
  const text = cleanText(value).toLowerCase();
  if (text.includes("wmass") || text.includes("western ma")) return "WMASS";
  if (text.includes("emass") || text.includes("eastern ma")) return "EMASS";
  if (text.includes("connecticut") || text === "ct") return "CT";
  if (text.includes("albany") || text.includes("new york") || text.includes("capital region")) return "Albany";
  return cleanText(value) || "Unknown";
}

function canonicalName(value) {
  const aliasMap = {
    dave: "david",
    matt: "matthew",
    jeff: "jeffrey",
    phil: "phillip",
    dan: "daniel",
  };

  const raw = cleanRepName(value)
    .toLowerCase()
    .replaceAll(".", " ")
    .replaceAll(",", " ")
    .replaceAll("-", " ")
    .replaceAll("_", " ")
    .replaceAll("'", " ");

  const parts = raw.split(" ").filter(Boolean);
  if (!parts.length) return "";
  const first = aliasMap[parts[0]] || parts[0];
  return [first, ...parts.slice(1)].join(" ").trim();
}

function normalizePercent(value, fallback = 0) {
  const n = safeNum(value);
  if (!n) return fallback;
  return n > 1 ? n / 100 : n;
}

function average(values, fallback = 0) {
  if (!values || !values.length) return fallback;
  return values.reduce((sum, value) => sum + safeNum(value), 0) / values.length;
}

function movementFromMetrics(rep) {
  if (safeNum(rep.nsli) >= 4500 && safeNum(rep.closePct) >= 0.26) return "Stable Elite";
  if (safeNum(rep.nsli) >= 4000 && safeNum(rep.closePct) >= 0.22) return "Breakout";
  if (safeNum(rep.closePct) < 0.18) return "Downtrend Risk";
  if (safeNum(rep.nsli) < 3500) return "Early Softening";
  return "Stable Core";
}

function createRepShell(rep, team = "Unknown") {
  return {
    rep: cleanRepName(rep),
    team: cleanTeamName(team),
    netVolume: 0,
    closePct: 0,
    netPct: 0.85,
    nsli: 4000,
    talkRatio: 0.52,
    scriptCompliance: 0.82,
    movement: "Imported",
    weekly: Array(16).fill(0),
    products: { Baths: 0, Windows: 0, Doors: 0, Roofing: 0, Other: 0 },
    _closeSamples: [],
    _netSamples: [],
    _nsliSamples: [],
    _talkSamples: [],
    _complianceSamples: [],
  };
}

function ensureRep(repMap, repName, team = "Unknown") {
  const cleaned = cleanRepName(repName);
  const key = canonicalName(cleaned);
  if (!key) return null;
  if (!repMap[key]) repMap[key] = createRepShell(cleaned, team);
  if (team && repMap[key].team === "Unknown") repMap[key].team = cleanTeamName(team);
  return repMap[key];
}

function coerceMaybeNumber(value) {
  if (value === null || value === undefined) return null;
  if (typeof value === "number") return Number.isFinite(value) ? value : null;
  const text = cleanText(value).split("$").join("").split(",").join("").split("%").join("");
  if (text === "") return null;
  const n = Number(text);
  return Number.isFinite(n) ? n : null;
}

function pushIfValue(target, value, transform = (x) => x) {
  const n = coerceMaybeNumber(value);
  if (n === null) return;
  target.push(transform(n));
}

function findColumnKey(row, matchers) {
  const keys = Object.keys(row || {});
  return keys.find((key) => matchers.some((matcher) => key.toLowerCase().includes(matcher)));
}

function finalizeRep(rep) {
  const weeklySum = rep.weekly.reduce((sum, value) => sum + safeNum(value), 0);
  const netVolume = weeklySum > 0 ? weeklySum : safeNum(rep.netVolume);
  const closePct = average(rep._closeSamples, rep.closePct || 0.22);
  const netPct = average(rep._netSamples, rep.netPct || 0.85);
  const nsli = average(rep._nsliSamples, rep.nsli || (netVolume > 0 ? netVolume / 40 : 4000));
  const talkRatio = average(rep._talkSamples, rep.talkRatio || 0.52);
  const scriptCompliance = average(rep._complianceSamples, rep.scriptCompliance || 0.82);

  return {
    rep: rep.rep,
    team: rep.team,
    netVolume,
    closePct,
    netPct,
    nsli,
    talkRatio,
    scriptCompliance,
    movement: rep.movement === "Imported" ? movementFromMetrics({ netVolume, closePct, nsli }) : rep.movement,
    weekly: rep.weekly.slice(0, 16),
    products: rep.products,
  };
}

function validRepName(name) {
  const cleaned = cleanRepName(name);
  if (!cleaned) return false;
  const lower = cleaned.toLowerCase();
  if (lower === "-") return false;
  if (lower === "total" || lower === "totals") return false;
  if (lower.includes("grand total")) return false;
  if (lower.includes("sub total") || lower.includes("subtotal")) return false;
  if (lower.includes("summary")) return false;
  return true;
}

function bonusForVolume(volume) {
  const monthlyVolume = safeNum(volume);
  const tier = bonusTiers.find((item) => monthlyVolume >= item.min && monthlyVolume <= item.max);
  if (!tier) return { amount: 0, label: "Below $125k", rate: 0 };
  return {
    amount: monthlyVolume * tier.rate,
    label: tier.label,
    rate: tier.rate,
  };
}

function clamp(value, min = 0, max = 1) {
  return Math.min(max, Math.max(min, safeNum(value)));
}

function estimatedLeadCounts(rep) {
  const issueCount = safeNum(rep.issueCount) || (safeNum(rep.nsli) > 0 ? safeNum(rep.netVolume) / safeNum(rep.nsli) : 0);
  const demoCount = safeNum(rep.demoCount) || issueCount * safeNum(rep.demoPct);
  const soldCount = safeNum(rep.soldCount) || demoCount * safeNum(rep.closePct);
  const netCount = safeNum(rep.netCount) || soldCount * safeNum(rep.netPct);
  const avgNetPerNetSale = safeNum(rep.avgNetPerNetSale) || (netCount > 0 ? safeNum(rep.netVolume) / netCount : 0);
  return { issueCount, demoCount, soldCount, netCount, avgNetPerNetSale };
}

function avgTicketFromRep(rep) {
  const { netCount } = estimatedLeadCounts(rep || {});
  return netCount > 0 ? safeNum(rep?.netVolume) / netCount : 0;
}

function avgTicketFromRepCollection(reps = []) {
  const totals = (reps || []).reduce(
    (acc, rep) => {
      const counts = estimatedLeadCounts(rep || {});
      acc.netVolume += safeNum(rep?.netVolume);
      acc.netCount += safeNum(counts.netCount);
      return acc;
    },
    { netVolume: 0, netCount: 0 }
  );

  return totals.netCount > 0 ? totals.netVolume / totals.netCount : 0;
}

function projectorMetricLabel(metric) {
  if (metric === "netVolume") return "Net Volume";
  if (metric === "demoPct") return "Demo %";
  if (metric === "closePct") return "Close %";
  if (metric === "netPct") return "Net %";
  if (metric === "avgTicket") return "Avg Ticket";
  return metric;
}

function percentileRank(value, values, higherIsBetter = true) {
  const cleanValues = (values || []).map((item) => safeNum(item)).filter((item) => Number.isFinite(item));
  if (!cleanValues.length) return 50;
  const sorted = [...cleanValues].sort((a, b) => a - b);
  const rank = sorted.filter((item) => item <= safeNum(value)).length / sorted.length;
  const pctRank = Math.round(rank * 100);
  return higherIsBetter ? pctRank : 100 - pctRank;
}

function dominantProduct(rep) {
  const products = Object.entries(rep.products || {}).sort((a, b) => safeNum(b[1]) - safeNum(a[1]));
  return products.length && safeNum(products[0][1]) > 0 ? products[0][0] : "core";
}

function coachingForRep(rep, peerReps = [], timeframe = "90") {
  const peers = (peerReps || []).filter((item) => item && item.rep !== rep.rep);
  const peerSet = peers.length ? peers : [rep];
  const avgClose = average(peerSet.map((item) => item.closePct), 0.24);
  const avgNet = average(peerSet.map((item) => item.netPct), 0.7);
  const avgNsli = average(peerSet.map((item) => item.nsli), 4000);
  const avgCompliance = average(peerSet.map((item) => item.scriptCompliance), 0.8);
  const topProduct = dominantProduct(rep);
  const closeRank = percentileRank(rep.closePct, peerSet.map((item) => item.closePct), true);
  const nsliRank = percentileRank(rep.nsli, peerSet.map((item) => item.nsli), true);
  const complianceRank = percentileRank(rep.scriptCompliance, peerSet.map((item) => item.scriptCompliance), true);

  const biggestGap = [
    { key: "close", score: Math.max(0, avgClose - safeNum(rep.closePct)) },
    { key: "nsli", score: Math.max(0, avgNsli - safeNum(rep.nsli)) / Math.max(1, avgNsli) },
    { key: "compliance", score: Math.max(0, avgCompliance - safeNum(rep.scriptCompliance)) },
    { key: "demo", score: Math.max(0, KPI_GOALS.demoPct - safeNum(rep.demoPct)) },
    { key: "net", score: Math.max(0, KPI_GOALS.netPct - safeNum(rep.netPct)) },
  ].sort((a, b) => b.score - a.score)[0];

  let gapGuidance = `Close rate is ${pct(rep.closePct)} versus a peer average of ${pct(avgClose)}. Improve the current-state to future-state contrast earlier in the appointment so the need feels more urgent before price discussion.`;
  if (rep.closePct >= avgClose && rep.netPct >= avgNet) {
    gapGuidance = `Close rate is ${pct(rep.closePct)} and net percent is ${pct(rep.netPct)}, both at or above peer averages. The next lever is better business-impact framing so wins in ${topProduct.toLowerCase()} drive more repeatable urgency.`;
  } else if (rep.closePct < 0.22) {
    gapGuidance = `Close rate is ${pct(rep.closePct)}, below the peer average of ${pct(avgClose)}. Spend more time defining the real cost of delay, what happens if the customer waits, and why the current state is no longer acceptable.`;
  }

  let cheeseGuidance = `Net percent is ${pct(rep.netPct)} versus a peer average of ${pct(avgNet)}. Coach this rep to adapt better when buyer conditions shift late in the process without defaulting to price concessions.`;
  if (rep.netPct >= avgNet && rep.movement.includes("Stable")) {
    cheeseGuidance = `Movement profile is ${rep.movement} with net percent at ${pct(rep.netPct)}. This rep is handling change reasonably well, so the focus should be on preserving momentum and not over-adjusting mid-appointment.`;
  } else if (rep.movement.includes("Downtrend") || rep.movement.includes("Softening")) {
    cheeseGuidance = `Movement profile is ${rep.movement}. The rep needs faster adjustment when the buyer goes uncertain late: re-anchor the problem, restate consequences, and simplify the next step instead of over-explaining.`;
  }

  let priorityGuidance = `NSLI is ${currency(rep.nsli)} versus the company goal of ${currency(KPI_GOALS.nsli)}. Protect premium positioning and make the recommendation feel like the logical solution, not just the expensive one.`;
  if (biggestGap.key === "close") {
    priorityGuidance = `Primary coaching priority: close rate. This rep sits around the ${closeRank}th percentile on close performance in the selected period. Focus the next review on deeper discovery, cost-of-delay language, and asking for commitment earlier.`;
  } else if (biggestGap.key === "nsli") {
    priorityGuidance = `Primary coaching priority: NSLI. At ${currency(rep.nsli)}, this rep trails the peer average of ${currency(avgNsli)}. The fastest lift is stronger premium anchoring and better value framing before lower-cost alternatives are discussed.`;
  } else if (biggestGap.key === "compliance") {
    priorityGuidance = `Primary coaching priority: script compliance. Compliance is ${pct(rep.scriptCompliance)} and sits around the ${complianceRank}th percentile. Tighten process discipline around discovery, urgency, and recommendation flow.`;
  } else if (biggestGap.key === "demo") {
    priorityGuidance = `Primary coaching priority: demo rate. Demo percent is ${pct(rep.demoPct)} against the company goal of ${pct(KPI_GOALS.demoPct)}. Focus on firmer set control, stronger appointment discipline, and reducing low-intent demos.`;
  } else if (biggestGap.key === "net") {
    priorityGuidance = `Primary coaching priority: net rate. Net percent is ${pct(rep.netPct)} against the company goal of ${pct(KPI_GOALS.netPct)}. Emphasize cleaner deal structure, better expectation setting, and fewer preventable cancels.`;
  }

  return [
    { framework: "Gap Selling", guidance: gapGuidance },
  ];
}

function buildMergedData(uploadData) {
  if (!uploadData || !uploadData.reps || !uploadData.reps.length) return demoReps;
  return uploadData.reps;
}

function buildWindowedDatasetFromWeekly(dataset, days) {
  const sourceWeeks = dataset?.weekLabels?.length ? dataset.weekLabels : demoWeeks;
  const sourceReps = buildMergedData(dataset);
  const weekCount = Math.max(1, Math.ceil(Math.max(1, safeNum(days)) / 7));
  const sliceStart = Math.max(0, sourceWeeks.length - weekCount);
  const weekLabels = sourceWeeks.slice(sliceStart);
  const bounds = getDatasetDateBoundsFromWeeks(weekLabels);

  const reps = sourceReps.map((rep) => {
    const weekly = (rep.weekly || []).slice(sliceStart, sliceStart + weekCount);
    const netVolume = weekly.reduce((sum, value) => sum + safeNum(value), 0);
    return {
      ...rep,
      weekly,
      netVolume,
    };
  });

  return {
    ...(dataset || {}),
    weekLabels,
    reps,
    minDate: bounds.minDate,
    maxDate: bounds.maxDate,
  };
}

function parseWeekLabelRange(label) {
  const parts = String(label || "").split(" - ");
  if (parts.length !== 2) return { start: null, end: null };
  return {
    start: parseUsDate(parts[0]),
    end: parseUsDate(parts[1]),
  };
}

function buildDateRangedDatasetFromWeekly(dataset, startIso, endIso) {
  const sourceWeeks = dataset?.weekLabels?.length ? dataset.weekLabels : demoWeeks;
  const sourceReps = buildMergedData(dataset);
  const startDate = startIso ? new Date(`${startIso}T00:00:00`) : null;
  const endDate = endIso ? new Date(`${endIso}T00:00:00`) : null;

  const overlappingWeeks = sourceWeeks
    .map((label, index) => ({ index, label, ...parseWeekLabelRange(label) }))
    .filter((item) => {
      if (!(item.start instanceof Date) || Number.isNaN(item.start?.getTime?.())) return false;
      if (!(item.end instanceof Date) || Number.isNaN(item.end?.getTime?.())) return false;
      if (startDate && item.end < startDate) return false;
      if (endDate && item.start > endDate) return false;
      return true;
    })
    .map((item) => {
      const overlapStart = startDate && item.start < startDate ? startDate : item.start;
      const overlapEnd = endDate && item.end > endDate ? endDate : item.end;
      const overlapMs = Math.max(0, overlapEnd.getTime() - overlapStart.getTime());
      const overlapDays = Math.max(1, Math.round(overlapMs / 86400000) + 1);
      const factor = Math.min(1, Math.max(0, overlapDays / 7));
      return {
        index: item.index,
        label: item.label,
        factor,
      };
    });

  const weekLabels = overlappingWeeks.map((item) => item.label);
  const reps = sourceReps.map((rep) => {
    const weekly = overlappingWeeks.map((item) => safeNum(rep.weekly?.[item.index]) * item.factor);
    const netVolume = weekly.reduce((sum, value) => sum + safeNum(value), 0);
    return {
      ...rep,
      weekly,
      netVolume,
    };
  });

  return {
    ...(dataset || {}),
    weekLabels,
    reps,
    minDate: startIso,
    maxDate: endIso,
  };
}

function dedupeAndNormalizeReps(inputReps, weekCount = 16) {
  const repMap = {};

  (inputReps || []).forEach((rep) => {
    const key = canonicalName(rep.rep);
    if (!key) return;

    if (!repMap[key]) {
      repMap[key] = {
        rep: cleanRepName(rep.rep),
        team: cleanTeamName(rep.team),
        movement: rep.movement || "Imported",
        weekly: Array.from({ length: weekCount }, (_, index) => safeNum(rep.weekly?.[index])),
        products: normalizeProductTotals(rep.products),
        _closeSamples: [],
        _netSamples: [],
        _nsliSamples: [],
        _talkSamples: [],
        _complianceSamples: [],
        _demoSamples: [],
        _marketingSamples: [],
      };
    } else {
      repMap[key].team = repMap[key].team === "Unknown" ? cleanTeamName(rep.team) : repMap[key].team;
      repMap[key].movement = repMap[key].movement === "Imported" ? (rep.movement || repMap[key].movement) : repMap[key].movement;
      repMap[key].weekly = repMap[key].weekly.map((value, index) => Math.max(safeNum(value), safeNum(rep.weekly?.[index])));
      ["Baths", "Windows", "Doors", "Roofing", "Siding"].forEach((product) => {
        repMap[key].products[product] = Math.max(safeNum(repMap[key].products[product]), safeNum(rep.products?.[product]));
      });
    }

    pushIfValue(repMap[key]._closeSamples, rep.closePct);
    pushIfValue(repMap[key]._netSamples, rep.netPct);
    pushIfValue(repMap[key]._nsliSamples, rep.nsli);
    pushIfValue(repMap[key]._talkSamples, rep.talkRatio);
    pushIfValue(repMap[key]._complianceSamples, rep.scriptCompliance);
    pushIfValue(repMap[key]._demoSamples, rep.demoPct);
    pushIfValue(repMap[key]._marketingSamples, rep.marketingScore);
  });

  return Object.values(repMap)
    .map((rep) => {
      const netVolume = rep.weekly.reduce((sum, value) => sum + safeNum(value), 0);
      const closePct = average(rep._closeSamples, 0);
      const netPct = average(rep._netSamples, 0);
      const nsli = average(rep._nsliSamples, netVolume > 0 ? netVolume / Math.max(1, rep.weekly.filter((value) => safeNum(value) > 0).length * 10) : 0);
      const talkRatio = average(rep._talkSamples, 0);
      const scriptCompliance = average(rep._complianceSamples, 0);
      const demoPct = average(rep._demoSamples, 0);
      const marketingScore = average(rep._marketingSamples, 0);

      return {
        rep: rep.rep,
        team: rep.team,
        netVolume,
        closePct,
        netPct,
        nsli,
        talkRatio,
        scriptCompliance,
        demoPct,
        marketingScore,
        movement: rep.movement === "Imported" ? movementFromMetrics({ netVolume, closePct, nsli }) : rep.movement,
        weekly: rep.weekly,
        products: rep.products,
      };
    })
    .sort((a, b) => b.netVolume - a.netVolume);
}

function tryParseWorkbook(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const names = workbook.SheetNames || [];
  const repMap = {};
  const weekLabels = [];

  function parseWeekOverWeekSalesMatrix(rows) {
    const headerIndex = rows.findIndex((row) => Array.isArray(row) && row.some((cell) => String(cell || "").toLowerCase().includes("leads issued")) && row.some((cell) => String(cell || "").toLowerCase().includes("net volume")));
    if (headerIndex === -1) return false;

    const header = rows[headerIndex] || [];
    const volumeIndex = header.findIndex((cell) => String(cell || "").toLowerCase().includes("net volume"));
    const closeIndex = header.findIndex((cell) => String(cell || "").toLowerCase().includes("close"));
    const netIndex = header.findIndex((cell) => String(cell || "").toLowerCase().includes("net %"));
    const nsliIndex = header.findIndex((cell) => String(cell || "").toLowerCase().includes("nsli"));

    let currentTeam = "Unknown";
    let weekIndex = -1;

    for (let i = headerIndex + 1; i < rows.length - 1; i += 1) {
      const label = cleanText(rows[i] && rows[i][0]);
      const metricRow = rows[i + 1] || [];
      if (!label) continue;

      if (label.startsWith("Appointment Date:")) {
        weekIndex += 1;
        const shortLabel = cleanText(label.replace("Appointment Date:", "").split("(")[0]);
        if (shortLabel && weekLabels.length <= weekIndex) weekLabels.push(shortLabel);
        continue;
      }

      if (label.includes("Sales Team:") && !label.includes("Sorted")) {
        currentTeam = cleanTeamName(label.replace("Sales Team:", ""));
        continue;
      }

      if (label.includes("Sales Rep:") && !label.includes("Sorted") && validRepName(label)) {
        const rep = ensureRep(repMap, label, currentTeam);
        if (!rep) continue;
        const volume = safeNum(metricRow[volumeIndex]);
        if (weekIndex >= 0 && weekIndex < rep.weekly.length) rep.weekly[weekIndex] = Math.max(rep.weekly[weekIndex], volume);
        rep.netVolume = Math.max(rep.netVolume, volume);
        pushIfValue(rep._closeSamples, metricRow[closeIndex], (n) => normalizePercent(n));
        pushIfValue(rep._netSamples, metricRow[netIndex], (n) => normalizePercent(n));
        pushIfValue(rep._nsliSamples, metricRow[nsliIndex]);
      }
    }

    return true;
  }

  function parseWeekOverWeekProductMatrix(rows) {
    const headerIndex = rows.findIndex((row) => Array.isArray(row) && row.some((cell) => String(cell || "").toLowerCase().includes("net amt") || String(cell || "").toLowerCase().includes("net volume")) && row.some((cell) => String(cell || "").toLowerCase().includes("set")));
    if (headerIndex === -1) return false;

    const header = rows[headerIndex] || [];
    const volumeIndex = header.findIndex((cell) => String(cell || "").toLowerCase().includes("net amt") || String(cell || "").toLowerCase().includes("net volume"));
    let currentProduct = "Other";

    for (let i = headerIndex + 1; i < rows.length - 1; i += 1) {
      const label = cleanText(rows[i] && rows[i][0]);
      const metricRow = rows[i + 1] || [];
      if (!label) continue;

      if (label.includes("Product:") && !label.includes("Sorted")) {
        currentProduct = cleanText(label.replace("Product:", "")) || "Other";
        continue;
      }

      if (label.includes("Sales Rep 1: Staff Name:") && !label.includes("Sorted") && validRepName(label)) {
        const rep = ensureRep(repMap, label, "Unknown");
        if (!rep) continue;
        const productName = ["Baths", "Windows", "Doors", "Roofing", "Siding", "Gutters"].includes(currentProduct) ? currentProduct : "Siding";
        const mappedProduct = ["Baths", "Windows", "Doors", "Roofing", "Siding"].includes(productName) ? productName : "Siding";
        rep.products[mappedProduct] = Math.max(rep.products[mappedProduct] || 0, safeNum(metricRow[volumeIndex]));
      }
    }

    return true;
  }

  function parseMarginReportMatrix(rows) {
    const headerIndex = rows.findIndex((row) => Array.isArray(row) && String(row[0] || "").toLowerCase().includes("sale: sale name") && String(row[2] || "").toLowerCase().includes("sold price"));
    if (headerIndex === -1) return false;

    let currentRep = null;
    for (let i = headerIndex + 1; i < rows.length; i += 1) {
      const label = cleanText(rows[i] && rows[i][0]);
      if (!label) continue;

      if (label.startsWith("Sales Rep 1:") && validRepName(label)) {
        currentRep = ensureRep(repMap, label, "Unknown");
        continue;
      }

      if (!currentRep) continue;
      if (label.startsWith("avg")) continue;
      if (!label.includes(":")) continue;

      const soldPrice = safeNum(rows[i] && rows[i][2]);
      currentRep.netVolume += soldPrice;

      const productsText = cleanText(label.split(":")[0]);
      productsText.split(";").map((part) => cleanText(part)).forEach((product) => {
        const mappedProduct = ["Baths", "Basic Baths", "Shower Conversion"].includes(product)
          ? "Baths"
          : ["Windows", "Doors", "Roofing", "Siding"].includes(product)
            ? product
            : "Siding";
        currentRep.products[mappedProduct] = safeNum(currentRep.products[mappedProduct]) + soldPrice;
      });

      pushIfValue(currentRep._netSamples, rows[i] && rows[i][3], (n) => normalizePercent(n));
    }

    return true;
  }

  function parseConversationRows(rows) {
    rows.forEach((row) => {
      const nameKey = findColumnKey(row, ["user name", "name"]);
      const talkKey = findColumnKey(row, ["talk ratio"]);
      const repName = cleanRepName(row[nameKey]);
      if (!validRepName(repName)) return;
      const rep = ensureRep(repMap, repName, row.Team || row.Teams || "Unknown");
      if (!rep) return;
      pushIfValue(rep._talkSamples, row[talkKey], (n) => normalizePercent(n));
    });
  }

  function parseVoiceRows(rows) {
    rows.forEach((row) => {
      const roleKey = findColumnKey(row, ["role"]);
      if (roleKey && String(row[roleKey] || "").toLowerCase() && !String(row[roleKey] || "").toLowerCase().includes("sales")) return;

      const nameKey = findColumnKey(row, ["name"]);
      const teamKey = findColumnKey(row, ["team"]);
      const complianceKey = findColumnKey(row, ["script compliance", "compliance"]);
      const repName = cleanRepName(row[nameKey]);
      if (!validRepName(repName)) return;
      const rep = ensureRep(repMap, repName, row[teamKey]);
      if (!rep) return;
      pushIfValue(rep._complianceSamples, row[complianceKey], (n) => normalizePercent(n));
    });
  }

  function parseFlatRepRows(rows) {
    rows.forEach((row) => {
      const nameKey = findColumnKey(row, ["sales rep", "rep", "staff name", "name"]);
      const teamKey = findColumnKey(row, ["team"]);
      const volumeKey = findColumnKey(row, ["net volume", "volume"]);
      const closeKey = findColumnKey(row, ["close"]);
      const netKey = findColumnKey(row, ["net %", "net percent"]);
      const nsliKey = findColumnKey(row, ["nsli"]);
      const talkKey = findColumnKey(row, ["talk ratio"]);
      const complianceKey = findColumnKey(row, ["script compliance", "compliance"]);

      const repName = cleanRepName(row[nameKey]);
      if (!validRepName(repName)) return;
      const rep = ensureRep(repMap, repName, row[teamKey]);
      if (!rep) return;

      rep.netVolume = Math.max(rep.netVolume, safeNum(row[volumeKey]));
      pushIfValue(rep._closeSamples, row[closeKey], (n) => normalizePercent(n));
      pushIfValue(rep._netSamples, row[netKey], (n) => normalizePercent(n));
      pushIfValue(rep._nsliSamples, row[nsliKey]);
      pushIfValue(rep._talkSamples, row[talkKey], (n) => normalizePercent(n));
      pushIfValue(rep._complianceSamples, row[complianceKey], (n) => normalizePercent(n));
    });
  }

  names.forEach((sheetName) => {
    const ws = workbook.Sheets[sheetName];
    if (!ws) return;

    const matrixRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
    const objectRows = XLSX.utils.sheet_to_json(ws, { defval: null });
    const lowerName = sheetName.toLowerCase();

    const parsedWeeklySales = parseWeekOverWeekSalesMatrix(matrixRows);
    const parsedWeeklyProducts = parseWeekOverWeekProductMatrix(matrixRows);
    const parsedMargin = lowerName.includes("margin") ? parseMarginReportMatrix(matrixRows) : false;

    if (lowerName.includes("conversation") || (objectRows[0] && Object.keys(objectRows[0]).some((key) => key.toLowerCase().includes("user name")))) {
      parseConversationRows(objectRows);
    }

    if (lowerName.includes("voice") || (objectRows[0] && Object.keys(objectRows[0]).some((key) => key.toLowerCase().includes("script compliance")))) {
      parseVoiceRows(objectRows);
    }

    if (!parsedWeeklySales && !parsedWeeklyProducts && !parsedMargin && objectRows.length) {
      parseFlatRepRows(objectRows);
    }
  });

  const effectiveWeekLabels = weekLabels.length ? weekLabels : demoWeeks;
  const bounds = getDatasetDateBoundsFromWeeks(effectiveWeekLabels);

  const reps = Object.values(repMap)
    .map((rep) => finalizeRep(rep))
    .filter((rep) => validRepName(rep.rep))
    .sort((a, b) => b.netVolume - a.netVolume);

  return {
    sheetNames: names,
    weekLabels: effectiveWeekLabels,
    reps,
    minDate: bounds.minDate,
    maxDate: bounds.maxDate,
  };
}

function parseUploadedArrayBuffer(arrayBuffer) {
  let parsed = null;
  const htmlText = new TextDecoder().decode(new Uint8Array(arrayBuffer));

  if (
    htmlText.includes("<!DOCTYPE HTML") ||
    htmlText.toLowerCase().includes("appointment date:") ||
    htmlText.toLowerCase().includes("product category:")
  ) {
    parsed = parseStandardHtmlReport(htmlText);
  }

  if (!parsed) {
    parsed = tryParseWorkbook(arrayBuffer);
  }

  return parsed;
}


function goalDeltaLabel(title, rawValue, goals = KPI_GOALS) {
  if (title === "NSLI") {
    const target = safeNum(goals?.nsli) || KPI_GOALS.nsli;
    const delta = safeNum(rawValue) - target;
    return `${delta >= 0 ? "+" : ""}${currency(delta)} vs goal`;
  }
  if (title === "Close %") {
    const target = safeNum(goals?.closePct) || KPI_GOALS.closePct;
    const delta = safeNum(rawValue) - target;
    return `${delta >= 0 ? "+" : ""}${pct(delta)} vs goal`;
  }
  if (title === "Demo %") {
    const target = safeNum(goals?.demoPct) || KPI_GOALS.demoPct;
    const delta = safeNum(rawValue) - target;
    return `${delta >= 0 ? "+" : ""}${pct(delta)} vs goal`;
  }
  if (title === "Net %") {
    const target = safeNum(goals?.netPct) || KPI_GOALS.netPct;
    const delta = safeNum(rawValue) - target;
    return `${delta >= 0 ? "+" : ""}${pct(delta)} vs goal`;
  }
  return "";
}

function metricAccentClass(title, rawValue, defaultAccent = "text-[var(--text-strong)]", goals = KPI_GOALS, bands = null) {
  const value = safeNum(rawValue);
  if (title === "NSLI") {
    const greenMin = safeNum(bands?.nsli?.greenMin) || safeNum(goals?.nsli) || KPI_GOALS.nsli;
    const yellowMin = safeNum(bands?.nsli?.yellowMin) || (greenMin * 0.925);
    if (value >= greenMin) return "text-[var(--kpi-good)]";
    if (value >= yellowMin) return "text-[var(--kpi-warn)]";
    return "text-[var(--kpi-bad)]";
  }
  if (title === "Close %") {
    const greenMin = safeNum(bands?.closePct?.greenMin) || safeNum(goals?.closePct) || KPI_GOALS.closePct;
    const yellowMin = safeNum(bands?.closePct?.yellowMin) || (greenMin * 0.9);
    if (value >= greenMin) return "text-[var(--kpi-good)]";
    if (value >= yellowMin) return "text-[var(--kpi-warn)]";
    return "text-[var(--kpi-bad)]";
  }
  if (title === "Demo %") {
    const greenMin = safeNum(bands?.demoPct?.greenMin) || safeNum(goals?.demoPct) || KPI_GOALS.demoPct;
    const yellowMin = safeNum(bands?.demoPct?.yellowMin) || (greenMin * 0.965);
    if (value >= greenMin) return "text-[var(--kpi-good)]";
    if (value >= yellowMin) return "text-[var(--kpi-warn)]";
    return "text-[var(--kpi-bad)]";
  }
  if (title === "Net %") {
    const greenMin = safeNum(bands?.netPct?.greenMin) || safeNum(goals?.netPct) || KPI_GOALS.netPct;
    const yellowMin = safeNum(bands?.netPct?.yellowMin) || (greenMin * 0.965);
    if (value >= greenMin) return "text-[var(--kpi-good)]";
    if (value >= yellowMin) return "text-[var(--kpi-warn)]";
    return "text-[var(--kpi-bad)]";
  }
  return defaultAccent;
}

function formatGoalEditorValue(value, inputType = "currency") {
  return inputType === "percent" ? `${Math.round(safeNum(value) * 100)}%` : currency(value);
}

function parseGoalEditorValue(rawValue, inputType = "currency") {
  const text = String(rawValue || "").trim();
  if (!text) return 0;
  if (inputType === "percent") {
    return safeNum(text.replace(/%/g, "")) / 100;
  }
  return parseMoney(text);
}

function repTierLetter(rep) {
  const closePct = safeNum(rep?.closePct);
  const nsli = safeNum(rep?.nsli);
  if (closePct >= 0.28 && nsli >= 4000) return "A";
  if (closePct >= 0.22 && nsli >= 3200) return "B";
  return "C";
}

function repHealthValue(rep) {
  const closeScore = clamp(safeNum(rep?.closePct) / KPI_GOALS.closePct, 0, 1.2);
  const netScore = clamp(safeNum(rep?.netPct) / KPI_GOALS.netPct, 0, 1.2);
  const nsliScore = clamp(safeNum(rep?.nsli) / KPI_GOALS.nsli, 0, 1.2);
  return ((closeScore * 0.35) + (netScore * 0.35) + (nsliScore * 0.3));
}

function teamTheme(teamName) {
  const key = cleanTeamName(teamName);
  if (key === "WMASS") return { header: "bg-emerald-500/10 text-emerald-300 border-emerald-500/30", pill: "text-emerald-300" };
  if (key === "EMASS") return { header: "bg-indigo-500/10 text-indigo-300 border-indigo-500/30", pill: "text-indigo-300" };
  if (key === "Albany") return { header: "bg-amber-500/10 text-amber-300 border-amber-500/30", pill: "text-amber-300" };
  if (key === "CT") return { header: "bg-sky-500/10 text-sky-300 border-sky-500/30", pill: "text-sky-300" };
  if (key === "Lobby") return { header: "bg-slate-500/10 text-slate-300 border-slate-500/30", pill: "text-slate-300" };
  return { header: "bg-violet-500/10 text-violet-300 border-violet-500/30", pill: "text-violet-300" };
}

const DEFAULT_MOVEMENT_CONFIG = {
  AA: { multiplier: 1.15, guidance: "Prioritize" },
  AB: { multiplier: 1.0, guidance: "Maintain" },
  AC: { multiplier: 0.9, guidance: "Maintain" },
  BA: { multiplier: 1.1, guidance: "Prioritize" },
  BB: { multiplier: 0.85, guidance: "Maintain" },
  BC: { multiplier: 0.75, guidance: "Cap / coach" },
  CA: { multiplier: 0.9, guidance: "Maintain" },
  CB: { multiplier: 0.85, guidance: "Maintain" },
  CC: { multiplier: 0.7, guidance: "Cap / coach" },
};

function classifyCloseTier(closePct, thresholds) {
  const value = safeNum(closePct);
  if (value >= safeNum(thresholds?.aMin)) return "A";
  if (value >= safeNum(thresholds?.bMin)) return "B";
  return "C";
}

function movementBucketMeta(bucketId) {
  const meta = {
    AA: { title: "Stable Elite", summary: "AA - held A in 90D and 30D", tone: "border-emerald-500/50", titleTone: "text-emerald-400" },
    AB: { title: "Stable", summary: "AB - A in 90D, slight dip to stable", tone: "border-emerald-500/40", titleTone: "text-emerald-400" },
    AC: { title: "Early Softening", summary: "AC - A in 90D but slipping to C", tone: "border-amber-500/50", titleTone: "text-amber-400" },
    BA: { title: "Breakout", summary: "BA - improved from B to A", tone: "border-sky-500/50", titleTone: "text-sky-400" },
    BB: { title: "Mediocrity", summary: "BB - held B in both windows", tone: "border-slate-500/50", titleTone: "text-slate-300" },
    BC: { title: "Downtrend Risk", summary: "BC - B in 90D, falling to C", tone: "border-rose-500/50", titleTone: "text-rose-400" },
    CA: { title: "Short-Term Spike", summary: "CA - C in 90D but spikes A in 30D", tone: "border-violet-500/50", titleTone: "text-violet-300" },
    CB: { title: "Early Recovery", summary: "CB - C in 90D, recovering to B", tone: "border-emerald-500/40", titleTone: "text-emerald-400" },
    CC: { title: "Chronic Underperformance", summary: "CC - held C in both windows", tone: "border-rose-500/50", titleTone: "text-rose-400" },
  };
  return meta[bucketId] || meta.BB;
}

function StatCard({ title, value, subvalue, secondaryValue = "", secondaryRawValue = null, tertiaryValue = "", tertiaryClassName = "text-[var(--kpi-goal)]", icon: Icon, accent = "text-[var(--text-strong)]", rawValue = null, goalNote = "", goals = KPI_GOALS, bands = null }) {
  const primaryLabel = subvalue || "";
  const secondaryLabel = secondaryValue ? secondaryValue.split(" ")[0] : "";
  const secondaryMetricValue = secondaryValue ? secondaryValue.slice(secondaryLabel.length).trim() : "";

  return (
    <Card className="border-[var(--border-strong)] bg-[var(--card-bg)]">
      <CardContent className="flex min-h-[112px] flex-col justify-between p-2.5">
        <div className="flex items-start justify-between gap-2">
          <div className="flex min-h-[87px] min-w-0 flex-1 flex-col justify-between">
            <div className="min-w-0">
              <div className="text-[9px] uppercase tracking-[0.14em] leading-tight text-[var(--kpi-title)]">{title}</div>
              <div
              className={`mt-1 font-bold leading-snug ${metricAccentClass(title, rawValue, accent, goals, bands)}`}
              style={{ fontSize: "27px", lineHeight: 1.15 }}
            >
                {primaryLabel ? <span style={{ fontSize: "18px" }}>{`${primaryLabel}:`}</span> : null}
              {primaryLabel ? " " : null}
              <span>{value}</span>
            </div>
              {secondaryValue ? (
                <div
                  className={`mt-[16px] font-bold leading-snug ${metricAccentClass(title, secondaryRawValue, "text-[var(--kpi-title)]", goals, bands)}`}
                  style={{ fontSize: "23px", lineHeight: 1.15 }}
                >
                  {secondaryLabel ? <span style={{ fontSize: "16px" }}>{`${secondaryLabel}:`}</span> : null}
                  {secondaryLabel ? " " : null}
                  <span>{secondaryMetricValue}</span>
                </div>
              ) : null}
            </div>
            <div className="min-w-0">
              {goalNote ? <div className="text-[15px] font-bold leading-snug text-[var(--kpi-goal)]">{goalNote}</div> : null}
              {tertiaryValue ? <div className={`text-[15px] font-bold leading-snug ${tertiaryClassName}`}>{tertiaryValue}</div> : null}
            </div>
          </div>
          <div className="shrink-0 rounded-lg border border-[var(--border)] bg-[var(--panel-bg)] p-1.5 text-[var(--text-soft)]">
            <Icon className="h-3.5 w-3.5" />
          </div>
        </div>
      </CardContent>
    </Card>
  );
}

function summarizeRawEntryMetrics(rawEntries = [], startIso = "", endIso = "", selectedGroup = "All Groups", selectedGroupMembers = new Set(), selectedRep = "All Reps", selectedProduct = "All Products", contextLabel = "") {
  const startDate = startIso ? new Date(`${startIso}T00:00:00`) : null;
  const endDate = endIso ? new Date(`${endIso}T00:00:00`) : null;
  const filteredEntries = (rawEntries || []).filter((entry) => {
    const date = parseUsDate(entry.date);
    if (!(date instanceof Date) || Number.isNaN(date.getTime())) return false;
    if (startDate && date < startDate) return false;
    if (endDate && date > endDate) return false;
    if (selectedGroup !== "All Groups" && !selectedGroupMembers.has(canonicalName(entry.rep))) return false;
    if (selectedRep !== "All Reps" && canonicalName(entry.rep) !== canonicalName(selectedRep)) return false;
    if (!productSelectionIncludes(selectedProduct, normalizeProductCategory(entry.product))) return false;
    return true;
  });

  const issueCount = filteredEntries.reduce((sum, entry) => sum + safeNum(entry.issueCount), 0);
  const demoCount = filteredEntries.reduce((sum, entry) => sum + safeNum(entry.demoCount), 0);
  const soldCount = filteredEntries.reduce((sum, entry) => sum + safeNum(entry.soldCount), 0);
  const netCount = filteredEntries.reduce((sum, entry) => sum + safeNum(entry.netCount), 0);
  const netVolume = filteredEntries.reduce((sum, entry) => sum + safeNum(entry.netVolume), 0);
  const grossVolume = filteredEntries.reduce((sum, entry) => sum + safeNum(entry.soldPriceAmount), 0);

  return {
    netVolume,
    grossVolume,
    closePct: issueCount > 0 ? soldCount / issueCount : 0,
    netPct: soldCount > 0 ? netCount / soldCount : 0,
    nsli: issueCount > 0 ? netVolume / issueCount : 0,
    demoPct: issueCount > 0 ? demoCount / issueCount : 0,
    avgTicket: netCount > 0 ? netVolume / netCount : 0,
    repCount: new Set(filteredEntries.map((entry) => canonicalName(entry.rep)).filter(Boolean)).size,
    contextLabel,
  };
}

function DatePickerField({ value, onChange, minDate, maxDate, placeholder = "Select date" }) {
  const inputId = useId();
  const inputRef = useRef(null);

  const openDatePicker = () => {
    const input = inputRef.current;
    if (!input) return;

    input.focus();
    if (typeof input.showPicker === "function") {
      input.showPicker();
      return;
    }
    input.click();
  };

  return (
    <div className="relative">
      <Input
        ref={inputRef}
        id={inputId}
        type="date"
        value={value || ""}
        min={minDate || undefined}
        max={maxDate || undefined}
        onChange={(event) => onChange(event.target.value)}
        className="h-11 border-[var(--border)] bg-[var(--panel-bg)] pr-14 text-[var(--text-strong)] [color-scheme:dark] [&::-webkit-calendar-picker-indicator]:opacity-0"
        aria-label={placeholder}
      />
      <button
        type="button"
        onClick={openDatePicker}
        aria-label={placeholder}
        className="absolute inset-y-0 right-0 flex w-14 cursor-pointer items-center justify-center rounded-r-md text-[var(--text-strong)] hover:bg-slate-800/70"
      >
        <CalendarDays className="h-5 w-5" />
      </button>
    </div>
  );
}

export default function SalesAnalyticsDashboardApp() {
  const [uploadMeta, setUploadMeta] = useState({ workbookName: DEFAULT_DATA_FILE_NAME, sheetNames: ["Embedded 30 Day", "Embedded 60 Day", "Embedded 90 Day"] });
  const [uploadData, setUploadData] = useState({
    timeframeSets: refreshedDatasets,
    reps: refreshedDatasets["90"].reps,
    weekLabels: refreshedDatasets["90"].weekLabels,
    sheetNames: ["Embedded 30 Day", "Embedded 60 Day", "Embedded 90 Day"],
    minDate: embeddedDemoBounds.minDate,
    maxDate: embeddedDemoBounds.maxDate,
  });
  const [selectedRep, setSelectedRep] = useState("All Reps");
  const [selectedGroup, setSelectedGroup] = useState("All Groups");
  const [customGroups, setCustomGroups] = useState({});
  const [groupNameInput, setGroupNameInput] = useState("");
  const [groupRepSearch, setGroupRepSearch] = useState("");
  const [groupDraftMembers, setGroupDraftMembers] = useState([]);
  const [repSearch, setRepSearch] = useState("");
  const [dateRange, setDateRange] = useState({ start: DEFAULT_DATE_START, end: DEFAULT_DATE_END });
  const [dashboardRangeEditorOpen, setDashboardRangeEditorOpen] = useState(false);
  const [dashboardDraftRange, setDashboardDraftRange] = useState({ start: DEFAULT_DATE_START, end: DEFAULT_DATE_END });
  const [scorecardDateRange, setScorecardDateRange] = useState({ start: DEFAULT_DATE_START, end: DEFAULT_DATE_END });
  const [scorecardRangeEditorOpen, setScorecardRangeEditorOpen] = useState(false);
  const [scorecardDraftRange, setScorecardDraftRange] = useState({ start: DEFAULT_DATE_START, end: DEFAULT_DATE_END });
  const [projectorAdjustments, setProjectorAdjustments] = useState({
    demoPct: 0,
    closePct: 0,
    netPct: 0,
    avgTicket: 0,
  });
  const [performanceTimeframe, setPerformanceTimeframe] = useState("90");
  const [projectorTimeframe, setProjectorTimeframe] = useState("90");
  const [projectorManualRange, setProjectorManualRange] = useState({ start: "", end: "" });
  const [projectorRangeEditorOpen, setProjectorRangeEditorOpen] = useState(false);
  const [projectorDraftRange, setProjectorDraftRange] = useState({ start: "", end: "" });
  const [selectedProduct, setSelectedProduct] = useState("All Products");
  const [weeklyTrendMetric, setWeeklyTrendMetric] = useState("netSales");
  const [weekOverWeekMetric, setWeekOverWeekMetric] = useState("nsli");
  const [productGraphMetric, setProductGraphMetric] = useState("nsli");
  const [volumeMetric, setVolumeMetric] = useState("net");
  const [analystPrompt, setAnalystPrompt] = useState("");
  const [analystMessages, setAnalystMessages] = useState([]);
  const [showFiltersPanel, setShowFiltersPanel] = useState(true);
  const [showTeamBuilderPanel, setShowTeamBuilderPanel] = useState(false);
  const [activeWorkspace, setActiveWorkspace] = useState("reviewStudio");
  const [activeDepartment, setActiveDepartment] = useState("Sales Department");
  const [goalTargets, setGoalTargets] = useState(DEFAULT_GOAL_TARGETS);
  const [kpiColorBands, setKpiColorBands] = useState(DEFAULT_KPI_COLOR_BANDS);
  const [annualVolumeBands, setAnnualVolumeBands] = useState(DEFAULT_ANNUAL_VOLUME_BANDS);
  const [tierThresholds, setTierThresholds] = useState(DEFAULT_TIER_THRESHOLDS);
  const [movementConfig, setMovementConfig] = useState(DEFAULT_MOVEMENT_CONFIG);
  const [isDarkMode, setIsDarkMode] = useState(false);
  const [isSidebarCollapsed, setIsSidebarCollapsed] = useState(false);

  const [commissionRate, setCommissionRate] = useState(0.1);

  useEffect(() => {
    let isCancelled = false;

    async function loadDefaultDataFile() {
      try {
        const parsed = parseStandardHtmlReport(DEFAULT_REPORT_HTML);

        if (!parsed || isCancelled) return;

        const initialRange = { start: DEFAULT_DATE_START, end: DEFAULT_DATE_END };

        setUploadMeta({ workbookName: DEFAULT_DATA_FILE_NAME, sheetNames: parsed.sheetNames || [] });
        setUploadData(parsed);
        setSelectedRep("All Reps");
        setRepSearch("");
        setDateRange(initialRange);
        setDashboardDraftRange(initialRange);
        setScorecardDateRange(initialRange);
        setScorecardDraftRange(initialRange);
        setProjectorManualRange(initialRange);
        setProjectorDraftRange(initialRange);
        setDashboardRangeEditorOpen(false);
        setScorecardRangeEditorOpen(false);
        setProjectorRangeEditorOpen(false);
      } catch (error) {
        console.error("Failed to load embedded default data file", error);
      }
    }

    loadDefaultDataFile();
    return () => {
      isCancelled = true;
    };
  }, []);

  const themeVars = useMemo(() => (
    isDarkMode
      ? {
          "--app-bg": "#171d26",
          "--app-text": "#e2e8f0",
          "--sidebar-bg": "#1f2733",
          "--sidebar-border": "#394557",
          "--header-bg": "rgba(27, 36, 48, 0.96)",
          "--card-bg": "#2d3642",
          "--panel-bg": "#212937",
          "--button-bg": "#283243",
          "--button-hover": "#313d52",
          "--button-active-bg": "#364253",
          "--border": "#4a5568",
          "--border-strong": "#465267",
          "--text-strong": "#f8fafc",
          "--text-soft": "#cbd5e1",
          "--sidebar-card-bg": "#2a3240",
          "--chart-label": "#cfd8e6",
          "--toggle-active-border": "#76a83a",
          "--toggle-active-bg": "#314926",
          "--toggle-active-text": "#86d448",
          "--toggle-active-ring": "#8fa6c7",
          "--selection-bg": "rgba(132, 204, 22, 0.12)",
          "--selection-text": "#bef264",
          "--upload-border": "#6f9f31",
          "--upload-bg": "#233822",
          "--upload-text": "#8fd13f",
          "--upload-hover-bg": "#284326",
          "--upload-hover-text": "#a4e04f",
          "--kpi-title": "#9aa7bb",
          "--kpi-sub": "#9aa7bb",
          "--kpi-goal": "#7f8ca1",
          "--kpi-good": "#4ade80",
          "--kpi-bad": "#f05252",
          "--kpi-warn": "#facc15",
          "--kpi-volume-primary": "#4ade80",
          "--kpi-volume-secondary": "#38bdf8",
        }
      : {
          "--app-bg": "#eef3f8",
          "--app-text": "#0f172a",
          "--sidebar-bg": "#e7edf4",
          "--sidebar-border": "#d7e0ea",
          "--header-bg": "rgba(231, 237, 244, 0.80)",
          "--card-bg": "#ffffff",
          "--panel-bg": "#f5f8fc",
          "--button-bg": "#edf2f7",
          "--button-hover": "#e3eaf3",
          "--button-active-bg": "#dce7f4",
          "--border": "#d0dae6",
          "--border-strong": "#c6d2df",
          "--text-strong": "#0f172a",
          "--text-soft": "#475569",
          "--sidebar-card-bg": "#ffffff",
          "--chart-label": "#334155",
          "--toggle-active-border": "#8ea4c3",
          "--toggle-active-bg": "#dce7f4",
          "--toggle-active-text": "#334155",
          "--toggle-active-ring": "#b8c8dc",
          "--selection-bg": "rgba(148, 163, 184, 0.14)",
          "--selection-text": "#334155",
          "--upload-border": "#8ea4c3",
          "--upload-bg": "#dce7f4",
          "--upload-text": "#334155",
          "--upload-hover-bg": "#e3eaf3",
          "--upload-hover-text": "#334155",
          "--kpi-title": "#475569",
          "--kpi-sub": "#334155",
          "--kpi-goal": "#475569",
          "--kpi-good": "#15803d",
          "--kpi-bad": "#ef4444",
          "--kpi-warn": "#facc15",
          "--kpi-volume-primary": "#166534",
          "--kpi-volume-secondary": "#1d4ed8"
        }
  ), [isDarkMode]);

  const datasetMinDate = uploadData?.minDate || embeddedDemoBounds.minDate;
  const datasetMaxDate = uploadData?.maxDate || embeddedDemoBounds.maxDate || toIsoDate(new Date());

  const scorecardRangeLabel = useMemo(() => {
    const start = scorecardDateRange.start || datasetMinDate;
    const end = scorecardDateRange.end || datasetMaxDate;
    if (start && end) return `${formatDisplayDate(start)} to ${formatDisplayDate(end)}`;
    return "Custom KPI period";
  }, [scorecardDateRange, datasetMinDate, datasetMaxDate]);

  const activeDataset = useMemo(() => {
    if (uploadData && uploadData.rawEntries && uploadData.rawEntries.length) {
      return aggregateRawEntriesToDataset(
        uploadData.rawEntries,
        dateRange.start || uploadData.minDate,
        dateRange.end || uploadData.maxDate
      );
    }

    const hasDynamicDateRange = Boolean(dateRange.start || dateRange.end);
    const baseWeeklyDataset = uploadData?.timeframeSets?.["90"] || (uploadData?.weekLabels?.length ? uploadData : { reps: demoReps, weekLabels: demoWeeks, minDate: embeddedDemoBounds.minDate, maxDate: embeddedDemoBounds.maxDate });

    if (hasDynamicDateRange) {
      return buildDateRangedDatasetFromWeekly(
        baseWeeklyDataset,
        dateRange.start || uploadData?.minDate || embeddedDemoBounds.minDate,
        dateRange.end || uploadData?.maxDate || embeddedDemoBounds.maxDate
      );
    }

    if (uploadData && uploadData.timeframeSets && uploadData.timeframeSets[performanceTimeframe]) {
      return uploadData.timeframeSets[performanceTimeframe];
    }
    if (uploadData && uploadData.reps) {
      return uploadData;
    }
    return { reps: demoReps, weekLabels: demoWeeks, minDate: embeddedDemoBounds.minDate, maxDate: embeddedDemoBounds.maxDate };
  }, [uploadData, performanceTimeframe, dateRange]);

  const scorecardDataset = useMemo(() => {
    if (uploadData && uploadData.rawEntries && uploadData.rawEntries.length) return null;

    const baseWeeklyDataset = uploadData?.timeframeSets?.["90"] || (uploadData?.weekLabels?.length ? uploadData : { reps: demoReps, weekLabels: demoWeeks, minDate: embeddedDemoBounds.minDate, maxDate: embeddedDemoBounds.maxDate });
    return buildDateRangedDatasetFromWeekly(
      baseWeeklyDataset,
      scorecardDateRange.start || uploadData?.minDate || embeddedDemoBounds.minDate,
      scorecardDateRange.end || uploadData?.maxDate || embeddedDemoBounds.maxDate
    );
  }, [uploadData, scorecardDateRange]);

  const activeWeeks = useMemo(() => {
    return activeDataset && activeDataset.weekLabels && activeDataset.weekLabels.length ? activeDataset.weekLabels : demoWeeks;
  }, [activeDataset]);

  const scorecardWeeks = useMemo(() => {
    return scorecardDataset && scorecardDataset.weekLabels && scorecardDataset.weekLabels.length ? scorecardDataset.weekLabels : activeWeeks;
  }, [scorecardDataset, activeWeeks]);

  const effectiveRangeLabel = useMemo(() => {
    if (dateRange.start && dateRange.end) {
      return `${formatDisplayDate(dateRange.start)} to ${formatDisplayDate(dateRange.end)}`;
    }
    if (activeDataset && activeDataset.minDate && activeDataset.maxDate) {
      return `${formatDisplayDate(activeDataset.minDate)} to ${formatDisplayDate(activeDataset.maxDate)}`;
    }
    if (uploadData && uploadData.minDate && uploadData.maxDate) {
      return `${formatDisplayDate(uploadData.minDate)} to ${formatDisplayDate(uploadData.maxDate)}`;
    }
    return `${performanceTimeframe}-day period`;
  }, [dateRange, activeDataset, uploadData, performanceTimeframe]);

  const reps = useMemo(() => {
    const baseReps = buildMergedData(activeDataset);

    if (uploadData && uploadData.rawEntries && uploadData.rawEntries.length) {
      return baseReps.map((rep) => ({
        ...rep,
        demoPct: safeNum(rep.demoPct),
        marketingScore: safeNum(rep.marketingScore),
      }));
    }

    const enriched = baseReps.map((rep) => ({
      ...rep,
      demoPct: rep.demoPct ?? getRepDemoPct(rep.rep, performanceTimeframe),
      marketingScore: safeNum(rep.marketingScore),
    }));

    return dedupeAndNormalizeReps(enriched, activeWeeks.length || 16);
  }, [activeDataset, uploadData, performanceTimeframe, activeWeeks]);

  const defaultTeamGroups = useMemo(() => {
    const grouped = Object.fromEntries(PRELOADED_GROUP_NAMES.map((name) => [name, []]));
    const seenByTeam = Object.fromEntries(PRELOADED_GROUP_NAMES.map((name) => [name, new Set()]));
    const rosterLookup = {};

    Object.entries(DEFAULT_TEAM_ROSTERS).forEach(([teamName, members]) => {
      members.forEach((member) => {
        rosterLookup[canonicalName(member)] = teamName;
      });
    });

    reps.forEach((rep) => {
      const canonicalRep = canonicalName(rep.rep);
      const rosterTeam = rosterLookup[canonicalRep];
      const fallbackTeam = cleanTeamName(rep.team);
      const teamName = rosterTeam || fallbackTeam;
      if (!grouped[teamName]) return;
      if (seenByTeam[teamName].has(canonicalRep)) return;
      grouped[teamName].push(rep.rep);
      seenByTeam[teamName].add(canonicalRep);
    });

    return grouped;
  }, [reps]);

  const effectiveGroups = useMemo(() => {
    const merged = { ...defaultTeamGroups };
    Object.entries(customGroups || {}).forEach(([name, members]) => {
      merged[name] = Array.from(new Set((members || []).filter(Boolean)));
    });
    return merged;
  }, [defaultTeamGroups, customGroups]);

  const manageableTeamNames = useMemo(() => {
    const ordered = [...PRELOADED_GROUP_NAMES, ...Object.keys(customGroups || {}).filter((name) => !PRELOADED_GROUP_NAMES.includes(name))];
    return Array.from(new Set([...ordered, "Lobby"]));
  }, [customGroups]);

  const teamManagementColumns = useMemo(() => {
    const assigned = new Set(Object.entries(effectiveGroups)
      .filter(([name]) => name !== "Lobby")
      .flatMap(([, members]) => members || []));

    return manageableTeamNames.map((teamName) => {
      const members = teamName === "Lobby"
        ? reps.filter((rep) => !assigned.has(rep.rep)).map((rep) => rep.rep)
        : (effectiveGroups[teamName] || []);

      const teamReps = members
        .map((repName) => reps.find((rep) => rep.rep === repName))
        .filter(Boolean)
        .sort((a, b) => safeNum(b.netVolume) - safeNum(a.netVolume));

      return {
        name: teamName,
        reps: teamReps,
        theme: teamTheme(teamName),
      };
    });
  }, [manageableTeamNames, effectiveGroups, reps]);

  const moveRepToTeam = (repName, targetTeam) => {
    const cleanedName = cleanRepName(repName);
    setCustomGroups(() => {
      const next = {};
      manageableTeamNames.filter((name) => name !== "Lobby").forEach((teamName) => {
        next[teamName] = (effectiveGroups[teamName] || []).filter((member) => member !== cleanedName);
      });
      if (targetTeam !== "Lobby") {
        next[targetTeam] = Array.from(new Set([...(next[targetTeam] || []), cleanedName]));
      }
      return next;
    });
  };

  const createManagedTeam = () => {
    const draftName = typeof window !== "undefined" ? window.prompt("Enter a new team name") : "";
    const name = cleanText(draftName);
    if (!name) return;
    setCustomGroups((current) => ({ ...current, [name]: current[name] || [] }));
  };

  const repOptions = useMemo(() => [{ value: "All Reps", label: "All Reps" }, ...reps.map((rep) => ({ value: rep.rep, label: rep.rep }))], [reps]);
  const groupOptions = useMemo(() => [{ value: "All Groups", label: "All Groups" }, ...Object.keys(effectiveGroups).map((name) => ({ value: name, label: name }))], [effectiveGroups]);
  const searchedRepOptions = useMemo(() => {
    const query = cleanText(repSearch).toLowerCase();
    if (!query) return repOptions;
    return repOptions.filter((option) => option.value === "All Reps" || option.label.toLowerCase().includes(query));
  }, [repOptions, repSearch]);
  const groupBuilderOptions = useMemo(() => {
    const query = cleanText(groupRepSearch).toLowerCase();
    return reps.filter((rep) => !query || rep.rep.toLowerCase().includes(query));
  }, [reps, groupRepSearch]);
  const selectedGroupMembers = useMemo(
    () => new Set(((effectiveGroups[selectedGroup] || []).map((name) => canonicalName(name))).filter(Boolean)),
    [effectiveGroups, selectedGroup]
  );

  const filteredReps = useMemo(() => {
    const groupScopedReps = selectedGroup === "All Groups"
      ? reps
      : reps.filter((rep) => (effectiveGroups[selectedGroup] || []).includes(rep.rep));
    return selectedRep === "All Reps" ? groupScopedReps : groupScopedReps.filter((rep) => rep.rep === selectedRep);
  }, [reps, selectedRep, selectedGroup, effectiveGroups]);
  const scorecardFilteredReps = useMemo(() => {
    if (!scorecardDataset || (uploadData && uploadData.rawEntries && uploadData.rawEntries.length)) return filteredReps;

    const scorecardDays = Math.max(1, scorecardWeeks.length || 1) * 7;
    const scorecardReps = dedupeAndNormalizeReps(
      buildMergedData(scorecardDataset).map((rep) => ({
        ...rep,
        demoPct: rep.demoPct ?? getEmbeddedRepBaseline(rep.rep, scorecardDays, uploadData?.timeframeSets).demoPct,
        marketingScore: safeNum(rep.marketingScore),
      })),
      scorecardWeeks.length || 16
    );
    const groupScopedReps = selectedGroup === "All Groups"
      ? scorecardReps
      : scorecardReps.filter((rep) => (effectiveGroups[selectedGroup] || []).includes(rep.rep));
    return selectedRep === "All Reps" ? groupScopedReps : groupScopedReps.filter((rep) => rep.rep === selectedRep);
  }, [scorecardDataset, uploadData, filteredReps, scorecardWeeks, selectedRep, selectedGroup, effectiveGroups]);
  const windowFilteredReps = useMemo(() => filteredReps, [filteredReps]);
  const selectedRepRecord = useMemo(() => {
    const baseRep = selectedRep === "All Reps" ? filteredReps[0] || reps[0] : reps.find((rep) => rep.rep === selectedRep);
    return baseRep || null;
  }, [selectedRep, filteredReps, reps]);

  const portfolio = useMemo(() => {
    const repCount = windowFilteredReps.length;
    const netVolume = windowFilteredReps.reduce((sum, rep) => sum + safeNum(rep.netVolume), 0);
    const avgClose = repCount ? windowFilteredReps.reduce((sum, rep) => sum + safeNum(rep.closePct), 0) / repCount : 0;
    const avgNet = repCount ? windowFilteredReps.reduce((sum, rep) => sum + safeNum(rep.netPct), 0) / repCount : 0;
    const avgNsli = repCount ? windowFilteredReps.reduce((sum, rep) => sum + safeNum(rep.nsli), 0) / repCount : 0;
    const avgDemo = repCount ? windowFilteredReps.reduce((sum, rep) => sum + safeNum(rep.demoPct), 0) / repCount : 0;
    return { repCount, netVolume, avgClose, avgNet, avgNsli, avgDemo };
  }, [windowFilteredReps]);

  const dashboardMetrics = useMemo(() => {
    if (uploadData && uploadData.rawEntries && uploadData.rawEntries.length) {
      return summarizeRawEntryMetrics(
        uploadData.rawEntries,
        dateRange.start,
        dateRange.end,
        selectedGroup,
        selectedGroupMembers,
        selectedRep,
        selectedProduct,
        effectiveRangeLabel
      );
    }

    const scoreRows = filteredReps.map((rep) => {
      const selectedVolume = selectedProduct === "All Products"
        ? safeNum(rep.netVolume)
        : safeNum(rep.netVolume) > 0
          ? productSelectionVolume(rep.products, selectedProduct)
          : 0;
      return {
        ...rep,
        selectedVolume,
      };
    }).filter((rep) => rep.selectedVolume > 0 || selectedProduct === "All Products");

    const totalVolume = scoreRows.reduce((sum, rep) => sum + safeNum(rep.selectedVolume), 0);
    const weightedAverage = (field) => {
      if (!scoreRows.length) return 0;
      if (totalVolume <= 0) return average(scoreRows.map((rep) => safeNum(rep[field])), 0);
      return scoreRows.reduce((sum, rep) => sum + safeNum(rep[field]) * (safeNum(rep.selectedVolume) / totalVolume), 0);
    };

    const estimatedNetCount = scoreRows.reduce((sum, rep) => {
      const counts = estimatedLeadCounts(rep);
      if (selectedProduct === "All Products") return sum + safeNum(counts.netCount);
      return sum + (safeNum(rep.netVolume) > 0 ? safeNum(counts.netCount) * (safeNum(rep.selectedVolume) / safeNum(rep.netVolume)) : 0);
    }, 0);
    const estimatedGrossVolume = weightedAverage("netPct") > 0 ? totalVolume / weightedAverage("netPct") : totalVolume;

    return {
      netVolume: totalVolume,
      grossVolume: estimatedGrossVolume,
      closePct: weightedAverage("closePct"),
      netPct: weightedAverage("netPct"),
      nsli: weightedAverage("nsli"),
      demoPct: weightedAverage("demoPct"),
      avgTicket: estimatedNetCount > 0 ? totalVolume / estimatedNetCount : 0,
      repCount: scoreRows.length,
      contextLabel: effectiveRangeLabel,
    };
  }, [uploadData, dateRange, selectedRep, selectedGroup, selectedGroupMembers, selectedProduct, filteredReps, effectiveRangeLabel]);

  const stickyHeaderMetrics = useMemo(() => {
    if (uploadData && uploadData.rawEntries && uploadData.rawEntries.length) {
      const anchorIso = uploadData.maxDate || datasetMaxDate || toIsoDate(new Date());
      const anchorDate = new Date(`${anchorIso}T00:00:00`);
      const yearStartIso = `${anchorDate.getFullYear()}-01-01`;
      const monthStartIso = `${anchorDate.getFullYear()}-${String(anchorDate.getMonth() + 1).padStart(2, "0")}-01`;

      return {
        anchorIso,
        ytd: summarizeRawEntryMetrics(uploadData.rawEntries, yearStartIso, anchorIso, selectedGroup, selectedGroupMembers, selectedRep, selectedProduct, "YTD"),
        mtd: summarizeRawEntryMetrics(uploadData.rawEntries, monthStartIso, anchorIso, selectedGroup, selectedGroupMembers, selectedRep, selectedProduct, "MTD"),
      };
    }

    return {
      anchorIso: datasetMaxDate,
      ytd: dashboardMetrics,
      mtd: dashboardMetrics,
    };
  }, [uploadData, datasetMaxDate, dashboardMetrics, selectedGroup, selectedGroupMembers, selectedRep, selectedProduct]);

  const stickyYtdPace = useMemo(() => {
    const anchorIso = stickyHeaderMetrics.anchorIso || datasetMaxDate || toIsoDate(new Date());
    const anchorDate = new Date(`${anchorIso}T00:00:00`);
    const yearStart = new Date(anchorDate.getFullYear(), 0, 1);
    const elapsedDays = Math.max(1, Math.floor((anchorDate.getTime() - yearStart.getTime()) / 86400000) + 1);
    const annualized = safeNum(stickyHeaderMetrics.ytd.netVolume) * (365 / elapsedDays);
    return {
      annualized,
      toneClass: annualized >= annualVolumeBands.greenMin ? "text-[var(--kpi-good)]" : annualized >= annualVolumeBands.yellowMin ? "text-[var(--kpi-warn)]" : "text-[var(--kpi-bad)]",
    };
  }, [stickyHeaderMetrics, datasetMaxDate, annualVolumeBands]);


  const productMixData = useMemo(() => {
    if (uploadData && uploadData.rawEntries && uploadData.rawEntries.length) {
      const startDate = dateRange.start ? new Date(`${dateRange.start}T00:00:00`) : null;
      const endDate = dateRange.end ? new Date(`${dateRange.end}T00:00:00`) : null;
      const grouped = {};

      uploadData.rawEntries.forEach((entry) => {
        const date = parseUsDate(entry.date);
        if (!(date instanceof Date) || Number.isNaN(date.getTime())) return;
        if (startDate && date < startDate) return;
        if (endDate && date > endDate) return;
        if (selectedGroup !== "All Groups" && !selectedGroupMembers.has(canonicalName(entry.rep))) return;
        if (selectedRep !== "All Reps" && canonicalName(entry.rep) !== canonicalName(selectedRep)) return;

        const product = normalizeProductCategory(entry.product);
        if (!productSelectionIncludes(selectedProduct, product)) return;
        if (!grouped[product]) grouped[product] = { netVolume: 0, grossVolume: 0, issueCount: 0, netCount: 0 };
        grouped[product].netVolume += safeNum(entry.netVolume);
        grouped[product].grossVolume += safeNum(entry.soldPriceAmount);
        grouped[product].issueCount += safeNum(entry.issueCount);
        grouped[product].netCount += safeNum(entry.netCount);
      });

      return Object.entries(grouped)
        .map(([name, totals]) => {
          const netVolume = safeNum(totals.netVolume);
          const grossVolume = safeNum(totals.grossVolume);
          const issueCount = safeNum(totals.issueCount);
          const netCount = safeNum(totals.netCount);
          const value = productGraphMetric === "nsli"
            ? (issueCount > 0 ? netVolume / issueCount : 0)
            : productGraphMetric === "netSales"
              ? netCount
              : (volumeMetric === "gross" ? grossVolume : netVolume);
          return { name, value };
        })
        .sort((a, b) => b.value - a.value);
    }

    const grouped = {};
    windowFilteredReps.forEach((rep) => {
      const avgTicket = Math.max(1, avgTicketFromRep(rep));
      Object.entries(rep.products || {}).forEach(([product, value]) => {
        if (!productSelectionIncludes(selectedProduct, product)) return;
        const netVolume = safeNum(value);
        const grossFactor = safeNum(rep.netPct) > 0 ? 1 / safeNum(rep.netPct) : 1;
        const grossVolume = netVolume * grossFactor;
        if (!grouped[product]) grouped[product] = { netVolume: 0, grossVolume: 0, weightedNsli: 0, estimatedNetSales: 0 };
        grouped[product].netVolume += netVolume;
        grouped[product].grossVolume += grossVolume;
        grouped[product].weightedNsli += netVolume * safeNum(rep.nsli);
        grouped[product].estimatedNetSales += netVolume / avgTicket;
      });
    });

    return Object.entries(grouped)
      .map(([name, totals]) => {
        const netVolume = safeNum(totals.netVolume);
        const grossVolume = safeNum(totals.grossVolume);
        const value = productGraphMetric === "nsli"
          ? (netVolume > 0 ? safeNum(totals.weightedNsli) / netVolume : 0)
          : productGraphMetric === "netSales"
            ? safeNum(totals.estimatedNetSales)
            : (volumeMetric === "gross" ? grossVolume : netVolume);
        return { name, value };
      })
      .sort((a, b) => b.value - a.value);
  }, [uploadData, dateRange, selectedRep, selectedGroup, selectedGroupMembers, selectedProduct, windowFilteredReps, productGraphMetric, volumeMetric]);

  const productGraphTitle = useMemo(() => {
    if (productGraphMetric === "nsli") return "Product NSLI";
    if (productGraphMetric === "netSales") return "Product Net Sales";
    return volumeMetric === "gross" ? "Product Gross Volume" : "Product Net Volume";
  }, [productGraphMetric, volumeMetric]);

  const productGraphAxisFormatter = useMemo(() => {
    if (productGraphMetric === "netSales") {
      return (value) => `${Math.round(safeNum(value))}`;
    }
    return (value) => `$${Math.round(safeNum(value) / 1000)}k`;
  }, [productGraphMetric]);

  const productGraphTooltipFormatter = useMemo(() => {
    if (productGraphMetric === "netSales") {
      return (value) => `${Math.round(safeNum(value))} net sales`;
    }
    return (value) => currency(value);
  }, [productGraphMetric]);

  const projectorDayCount = useMemo(() => {
    if (projectorTimeframe !== "manual") {
      return safeNum(projectorTimeframe) || 90;
    }
    const manualStart = projectorManualRange.start || dateRange.start || uploadData?.minDate || "";
    const manualEnd = projectorManualRange.end || dateRange.end || uploadData?.maxDate || toIsoDate(new Date());
    if (!manualStart || !manualEnd) return 30;
    const startDate = new Date(`${manualStart}T00:00:00`);
    const endDate = new Date(`${manualEnd}T00:00:00`);
    const diff = Math.round((endDate.getTime() - startDate.getTime()) / 86400000) + 1;
    return Math.max(1, diff);
  }, [projectorTimeframe, projectorManualRange, dateRange, uploadData]);

  const projectorDataset = useMemo(() => {
    const manualStart = projectorManualRange.start || dateRange.start || uploadData?.minDate || embeddedDemoBounds.minDate;
    const manualEnd = projectorManualRange.end || dateRange.end || uploadData?.maxDate || embeddedDemoBounds.maxDate || toIsoDate(new Date());
    const anchorEnd = dateRange.end || uploadData?.maxDate || embeddedDemoBounds.maxDate || toIsoDate(new Date());
    const baseWeeklyDataset = uploadData?.timeframeSets?.["90"] || (uploadData?.weekLabels?.length ? uploadData : { reps: demoReps, weekLabels: demoWeeks, minDate: embeddedDemoBounds.minDate, maxDate: embeddedDemoBounds.maxDate });

    if (uploadData && uploadData.rawEntries && uploadData.rawEntries.length) {
      if (projectorTimeframe === "manual") {
        return aggregateRawEntriesToDataset(uploadData.rawEntries, manualStart, manualEnd);
      }
      const endDate = new Date(`${anchorEnd}T00:00:00`);
      const startDate = new Date(endDate);
      startDate.setDate(startDate.getDate() - (projectorDayCount - 1));
      return aggregateRawEntriesToDataset(uploadData.rawEntries, toIsoDate(startDate), toIsoDate(endDate));
    }

    if (projectorTimeframe === "manual") {
      return buildDateRangedDatasetFromWeekly(baseWeeklyDataset, manualStart, manualEnd);
    }

    const endDate = new Date(`${anchorEnd}T00:00:00`);
    const startDate = new Date(endDate);
    startDate.setDate(startDate.getDate() - (projectorDayCount - 1));
    return buildDateRangedDatasetFromWeekly(baseWeeklyDataset, toIsoDate(startDate), toIsoDate(endDate));
  }, [uploadData, dateRange, projectorTimeframe, projectorDayCount, projectorManualRange]);

  const projectorRangeLabel = useMemo(() => {
    if (projectorDataset?.minDate && projectorDataset?.maxDate) {
      return `${formatDisplayDate(projectorDataset.minDate)} to ${formatDisplayDate(projectorDataset.maxDate)}`;
    }
    return projectorTimeframe === "manual" ? "Manual range" : `${projectorTimeframe}-day period`;
  }, [projectorDataset, projectorTimeframe]);

  const projectorReps = useMemo(() => {
    const baseReps = buildMergedData(projectorDataset);

    if (uploadData && uploadData.rawEntries && uploadData.rawEntries.length) {
      return baseReps.map((rep) => ({
        ...rep,
        demoPct: safeNum(rep.demoPct),
      }));
    }

    const enriched = baseReps.map((rep) => {
      const baseline = getEmbeddedRepBaseline(rep.rep, projectorDayCount, uploadData?.timeframeSets || refreshedDatasets);
      return {
        ...rep,
        closePct: baseline.closePct || safeNum(rep.closePct),
        netPct: baseline.netPct || safeNum(rep.netPct),
        nsli: baseline.nsli || safeNum(rep.nsli),
        demoPct: baseline.demoPct || safeNum(rep.demoPct),
      };
    });

    return dedupeAndNormalizeReps(enriched, projectorDataset?.weekLabels?.length || 16);
  }, [projectorDataset, uploadData, projectorDayCount]);

  const projectorFilteredReps = useMemo(() => {
    const groupScopedReps = selectedGroup === "All Groups"
      ? projectorReps
      : projectorReps.filter((rep) => (effectiveGroups[selectedGroup] || []).includes(rep.rep));
    return selectedRep === "All Reps" ? groupScopedReps : groupScopedReps.filter((rep) => rep.rep === selectedRep);
  }, [projectorReps, selectedRep, selectedGroup, effectiveGroups]);

  const projectorSelectedRepRecord = useMemo(() => {
    if (!projectorFilteredReps.length) return null;
    return projectorFilteredReps[0] || null;
  }, [projectorFilteredReps]);

  const aggregateProjectionDelta = useMemo(() => {
    const values = Object.values(projectorAdjustments || {}).map((value) => safeNum(value));
    return values.length ? values.reduce((sum, value) => sum + value, 0) / values.length : 0;
  }, [projectorAdjustments]);

  const projectionTone = useMemo(() => {
    if (aggregateProjectionDelta > 0) {
      return {
        panel: "border-lime-500/40 bg-lime-500/10",
        label: "text-lime-200",
        value: "text-lime-300",
        cardBorder: "border-lime-500/30",
      };
    }
    if (aggregateProjectionDelta < 0) {
      return {
        panel: "border-orange-500/40 bg-orange-500/10",
        label: "text-orange-200",
        value: "text-orange-300",
        cardBorder: "border-orange-500/30",
      };
    }
    return {
      panel: "border-[var(--border)] bg-[var(--panel-bg)]/50",
      label: "text-[var(--text-soft)]",
      value: "text-[var(--text-strong)]",
      cardBorder: "border-[var(--border)]",
    };
  }, [aggregateProjectionDelta]);

  const projected = useMemo(() => {
    if (!projectorSelectedRepRecord) return null;

    const baselineVolume = safeNum(projectorSelectedRepRecord.netVolume);
    const monthsEquivalent = Math.max(1, projectorDayCount / 30);
    const { issueCount, avgNetPerNetSale } = estimatedLeadCounts(projectorSelectedRepRecord);

    const baseDemoPct = clamp(safeNum(projectorSelectedRepRecord.demoPct));
    const baseClosePct = clamp(safeNum(projectorSelectedRepRecord.closePct));
    const baseNetPct = clamp(safeNum(projectorSelectedRepRecord.netPct));
    const baseAvgTicket = Math.max(0, safeNum(avgNetPerNetSale));

    const projectedDemoPct = clamp(baseDemoPct * (1 + safeNum(projectorAdjustments.demoPct) / 100));
    const projectedClosePct = clamp(baseClosePct * (1 + safeNum(projectorAdjustments.closePct) / 100));
    const projectedNetPct = clamp(baseNetPct * (1 + safeNum(projectorAdjustments.netPct) / 100));
    const projectedAvgTicket = Math.max(0, baseAvgTicket * (1 + safeNum(projectorAdjustments.avgTicket) / 100));

    const demoFactor = baseDemoPct > 0 ? projectedDemoPct / baseDemoPct : 1;
    const closeFactor = baseClosePct > 0 ? projectedClosePct / baseClosePct : 1;
    const netFactor = baseNetPct > 0 ? projectedNetPct / baseNetPct : 1;
    const avgTicketFactor = baseAvgTicket > 0 ? projectedAvgTicket / baseAvgTicket : 1;
    const combinedLift = demoFactor * closeFactor * netFactor * avgTicketFactor;
    const projectedVolume = baselineVolume * combinedLift;

    const monthlyPace = projectedVolume / monthsEquivalent;
    const monthlyCommission = monthlyPace * commissionRate;
    const periodCommission = projectedVolume * commissionRate;
    const monthlyBonus = bonusForVolume(monthlyPace);
    const periodBonus = monthlyBonus.amount * monthsEquivalent;
    const monthlyIncome = monthlyCommission + monthlyBonus.amount;

    return {
      baseRep: projectorSelectedRepRecord.rep,
      baselineVolume,
      projectedVolume,
      monthsEquivalent,
      monthlyPace,
      monthlyCommission,
      monthlyIncome,
      periodCommission,
      monthlyBonusAmount: monthlyBonus.amount,
      monthlyBonusLabel: monthlyBonus.label,
      periodBonus,
      totalIncome: periodCommission + periodBonus,
      issueCount,
      baselineKpis: {
        demoPct: baseDemoPct,
        closePct: baseClosePct,
        netPct: baseNetPct,
        avgTicket: baseAvgTicket,
      },
      projectedKpis: {
        demoPct: projectedDemoPct,
        closePct: projectedClosePct,
        netPct: projectedNetPct,
        avgTicket: projectedAvgTicket,
      },
      combinedLift,
    };
  }, [projectorSelectedRepRecord, projectorAdjustments, commissionRate, projectorTimeframe, projectorDayCount]);

  const performanceSeries = useMemo(() => {
    if (uploadData && uploadData.rawEntries && uploadData.rawEntries.length) {
      const startDate = dateRange.start ? new Date(`${dateRange.start}T00:00:00`) : null;
      const endDate = dateRange.end ? new Date(`${dateRange.end}T00:00:00`) : null;
      const filteredEntries = uploadData.rawEntries.filter((entry) => {
        const date = parseUsDate(entry.date);
        if (!(date instanceof Date) || Number.isNaN(date.getTime())) return false;
        if (startDate && date < startDate) return false;
        if (endDate && date > endDate) return false;
        if (selectedGroup !== "All Groups" && !selectedGroupMembers.has(canonicalName(entry.rep))) return false;
        if (selectedRep !== "All Reps" && canonicalName(entry.rep) !== canonicalName(selectedRep)) return false;
        if (!productSelectionIncludes(selectedProduct, normalizeProductCategory(entry.product))) return false;
        return true;
      });

      const sums = {};
      filteredEntries.forEach((entry) => {
        const key = formatWeekRangeFromDate(parseUsDate(entry.date));
        sums[key] = (sums[key] || 0) + safeNum(entry.netVolume);
      });

      return activeWeeks.map((week) => ({ label: week, metricValue: safeNum(sums[week]) }));
    }

    const base = selectedRep === "All Reps" ? filteredReps : [selectedRepRecord].filter(Boolean);
    return activeWeeks.map((week, sourceIndex) => {
      const metricValue = base.reduce((sum, rep) => {
        const weeklyValue = safeNum(rep.weekly[sourceIndex]);
        if (selectedProduct === "All Products") return sum + weeklyValue;
        return sum + weeklyValue * (safeNum(rep.products[selectedProduct]) / Math.max(1, rep.netVolume));
      }, 0);
      return { label: week, metricValue };
    });
  }, [uploadData, dateRange, selectedRep, selectedGroup, selectedGroupMembers, selectedProduct, filteredReps, selectedRepRecord, activeWeeks]);

  const grossPerformanceSeries = useMemo(() => {
    if (uploadData && uploadData.rawEntries && uploadData.rawEntries.length) {
      const startDate = dateRange.start ? new Date(`${dateRange.start}T00:00:00`) : null;
      const endDate = dateRange.end ? new Date(`${dateRange.end}T00:00:00`) : null;
      const filteredEntries = uploadData.rawEntries.filter((entry) => {
        const date = parseUsDate(entry.date);
        if (!(date instanceof Date) || Number.isNaN(date.getTime())) return false;
        if (startDate && date < startDate) return false;
        if (endDate && date > endDate) return false;
        if (selectedGroup !== "All Groups" && !selectedGroupMembers.has(canonicalName(entry.rep))) return false;
        if (selectedRep !== "All Reps" && canonicalName(entry.rep) !== canonicalName(selectedRep)) return false;
        if (!productSelectionIncludes(selectedProduct, normalizeProductCategory(entry.product))) return false;
        return true;
      });

      const sums = {};
      filteredEntries.forEach((entry) => {
        const key = formatWeekRangeFromDate(parseUsDate(entry.date));
        sums[key] = (sums[key] || 0) + safeNum(entry.soldPriceAmount);
      });

      return activeWeeks.map((week) => ({ label: week, metricValue: safeNum(sums[week]) }));
    }

    const base = selectedRep === "All Reps" ? filteredReps : [selectedRepRecord].filter(Boolean);
    return activeWeeks.map((week, sourceIndex) => {
      const metricValue = base.reduce((sum, rep) => {
        const weeklyNetValue = safeNum(rep.weekly[sourceIndex]);
        const scopedNetValue = selectedProduct === "All Products"
          ? weeklyNetValue
          : weeklyNetValue * (safeNum(rep.products[selectedProduct]) / Math.max(1, rep.netVolume));
        const grossFactor = safeNum(rep.netPct) > 0 ? 1 / safeNum(rep.netPct) : 1;
        return sum + scopedNetValue * grossFactor;
      }, 0);
      return { label: week, metricValue };
    });
  }, [uploadData, dateRange, selectedRep, selectedGroup, selectedGroupMembers, selectedProduct, filteredReps, selectedRepRecord, activeWeeks]);

  const weeklyTrendSeries = useMemo(() => {
    if (weeklyTrendMetric === "volume") return volumeMetric === "gross" ? grossPerformanceSeries : performanceSeries;

    if (uploadData && uploadData.rawEntries && uploadData.rawEntries.length) {
      const startDate = dateRange.start ? new Date(`${dateRange.start}T00:00:00`) : null;
      const endDate = dateRange.end ? new Date(`${dateRange.end}T00:00:00`) : null;
      const filteredEntries = uploadData.rawEntries.filter((entry) => {
        const date = parseUsDate(entry.date);
        if (!(date instanceof Date) || Number.isNaN(date.getTime())) return false;
        if (startDate && date < startDate) return false;
        if (endDate && date > endDate) return false;
        if (selectedGroup !== "All Groups" && !selectedGroupMembers.has(canonicalName(entry.rep))) return false;
        if (selectedRep !== "All Reps" && canonicalName(entry.rep) !== canonicalName(selectedRep)) return false;
        if (!productSelectionIncludes(selectedProduct, normalizeProductCategory(entry.product))) return false;
        return true;
      });
      const sums = {};
      filteredEntries.forEach((entry) => {
        const key = formatWeekRangeFromDate(parseUsDate(entry.date));
        sums[key] = (sums[key] || 0) + safeNum(entry.netCount);
      });

      return activeWeeks.map((week) => ({ label: week, metricValue: safeNum(sums[week]) }));
    }

    const base = selectedRep === "All Reps" ? filteredReps : [selectedRepRecord].filter(Boolean);
    return activeWeeks.map((week, sourceIndex) => {
      const metricValue = base.reduce((sum, rep) => {
        const weeklyValue = safeNum(rep.weekly[sourceIndex]);
        const scopedWeeklyValue = selectedProduct === "All Products"
          ? weeklyValue
          : weeklyValue * (productSelectionVolume(rep.products, selectedProduct) / Math.max(1, rep.netVolume));
        const avgTicket = Math.max(1, avgTicketFromRep(rep));
        return sum + (scopedWeeklyValue / avgTicket);
      }, 0);
      return { label: week, metricValue };
    });
  }, [weeklyTrendMetric, volumeMetric, performanceSeries, grossPerformanceSeries, uploadData, dateRange, selectedRep, selectedGroup, selectedGroupMembers, selectedProduct, filteredReps, selectedRepRecord, activeWeeks]);

  const weeklyTrendDisplaySeries = useMemo(() => {
    if (weeklyTrendMetric !== "volume") return weeklyTrendSeries;
    return activeWeeks.map((week, index) => ({
      label: week,
      netVolume: safeNum(performanceSeries[index]?.metricValue),
      grossVolume: safeNum(grossPerformanceSeries[index]?.metricValue),
      metricValue: safeNum(performanceSeries[index]?.metricValue),
    }));
  }, [weeklyTrendMetric, weeklyTrendSeries, activeWeeks, performanceSeries, grossPerformanceSeries]);

  const isWeeklyTrendNetSales = weeklyTrendMetric === "netSales";
  const weeklyTrendAxisFormatter = (value) => isWeeklyTrendNetSales
    ? `${Math.round(safeNum(value))}`
    : `$${Math.round(safeNum(value) / 1000)}k`;

  const weeklyTrendTooltipFormatter = (value) => isWeeklyTrendNetSales
    ? `${Math.round(safeNum(value))} net sales`
    : currency(value);

  const weekOverWeekBaseSeries = useMemo(() => {
    if (weekOverWeekMetric === "volume") return volumeMetric === "gross" ? grossPerformanceSeries : performanceSeries;
    if (weekOverWeekMetric === "netSales") return weeklyTrendSeries;

    if (uploadData && uploadData.rawEntries && uploadData.rawEntries.length) {
      const startDate = dateRange.start ? new Date(`${dateRange.start}T00:00:00`) : null;
      const endDate = dateRange.end ? new Date(`${dateRange.end}T00:00:00`) : null;
      const filteredEntries = uploadData.rawEntries.filter((entry) => {
        const date = parseUsDate(entry.date);
        if (!(date instanceof Date) || Number.isNaN(date.getTime())) return false;
        if (startDate && date < startDate) return false;
        if (endDate && date > endDate) return false;
        if (selectedGroup !== "All Groups" && !selectedGroupMembers.has(canonicalName(entry.rep))) return false;
        if (selectedRep !== "All Reps" && canonicalName(entry.rep) !== canonicalName(selectedRep)) return false;
        if (!productSelectionIncludes(selectedProduct, normalizeProductCategory(entry.product))) return false;
        return true;
      });

      const sums = {};
      filteredEntries.forEach((entry) => {
        const key = formatWeekRangeFromDate(parseUsDate(entry.date));
        if (!sums[key]) sums[key] = { netVolume: 0, issueCount: 0 };
        sums[key].netVolume += safeNum(entry.netVolume);
        sums[key].issueCount += safeNum(entry.issueCount);
      });

      return activeWeeks.map((week) => {
        const totals = sums[week] || { netVolume: 0, issueCount: 0 };
        return {
          label: week,
          metricValue: safeNum(totals.issueCount) > 0 ? safeNum(totals.netVolume) / safeNum(totals.issueCount) : 0,
        };
      });
    }

    const base = selectedRep === "All Reps" ? filteredReps : [selectedRepRecord].filter(Boolean);
    return activeWeeks.map((week, sourceIndex) => {
      const totals = base.reduce(
        (acc, rep) => {
          const weeklyValue = safeNum(rep.weekly[sourceIndex]);
          const scopedWeeklyValue = selectedProduct === "All Products"
            ? weeklyValue
            : weeklyValue * (safeNum(rep.products[selectedProduct]) / Math.max(1, rep.netVolume));
          acc.netVolume += scopedWeeklyValue;
          acc.issueCount += safeNum(rep.nsli) > 0 ? scopedWeeklyValue / safeNum(rep.nsli) : 0;
          return acc;
        },
        { netVolume: 0, issueCount: 0 }
      );

      return {
        label: week,
        metricValue: safeNum(totals.issueCount) > 0 ? safeNum(totals.netVolume) / safeNum(totals.issueCount) : 0,
      };
    });
  }, [weekOverWeekMetric, volumeMetric, performanceSeries, grossPerformanceSeries, weeklyTrendSeries, uploadData, dateRange, selectedRep, selectedGroup, selectedGroupMembers, selectedProduct, filteredReps, selectedRepRecord, activeWeeks]);

  const weekChangeData = useMemo(() => {
    return weekOverWeekBaseSeries.map((row, index) => ({
      label: row.label,
      actualValue: row.metricValue,
      delta: index === 0 ? 0 : row.metricValue - weekOverWeekBaseSeries[index - 1].metricValue,
    }));
  }, [weekOverWeekBaseSeries]);

  const weekOverWeekAxisFormatter = useMemo(() => {
    if (weekOverWeekMetric === "netSales") {
      return (value) => `${Math.round(safeNum(value))}`;
    }
    if (weekOverWeekMetric === "nsli") {
      return (value) => currency(value);
    }
    return (value) => `$${Math.round(safeNum(value) / 1000)}k`;
  }, [weekOverWeekMetric]);

  const weekOverWeekTooltipFormatter = useMemo(() => {
    if (weekOverWeekMetric === "netSales") {
      return (value) => `${Math.round(safeNum(value))} net sales`;
    }
    return (value) => currency(value);
  }, [weekOverWeekMetric]);
  const performanceScorecard = useMemo(() => {
    if (uploadData && uploadData.rawEntries && uploadData.rawEntries.length) {
      const startDate = scorecardDateRange.start ? new Date(`${scorecardDateRange.start}T00:00:00`) : null;
      const endDate = scorecardDateRange.end ? new Date(`${scorecardDateRange.end}T00:00:00`) : null;
      const filteredEntries = uploadData.rawEntries.filter((entry) => {
        const date = parseUsDate(entry.date);
        if (!(date instanceof Date) || Number.isNaN(date.getTime())) return false;
        if (startDate && date < startDate) return false;
        if (endDate && date > endDate) return false;
        if (selectedGroup !== "All Groups" && !selectedGroupMembers.has(canonicalName(entry.rep))) return false;
        if (selectedRep !== "All Reps" && canonicalName(entry.rep) !== canonicalName(selectedRep)) return false;
        if (!productSelectionIncludes(selectedProduct, normalizeProductCategory(entry.product))) return false;
        return true;
      });

      const issueCount = filteredEntries.reduce((sum, entry) => sum + safeNum(entry.issueCount), 0);
      const demoCount = filteredEntries.reduce((sum, entry) => sum + safeNum(entry.demoCount), 0);
      const soldCount = filteredEntries.reduce((sum, entry) => sum + safeNum(entry.soldCount), 0);
      const soldPriceAmount = filteredEntries.reduce((sum, entry) => sum + safeNum(entry.soldPriceAmount), 0);
      const netCount = filteredEntries.reduce((sum, entry) => sum + safeNum(entry.netCount), 0);
      const netVolume = filteredEntries.reduce((sum, entry) => sum + safeNum(entry.netVolume), 0);
      const workingAmount = filteredEntries.reduce((sum, entry) => sum + safeNum(entry.workingAmount), 0);
      const repCount = new Set(filteredEntries.map((entry) => canonicalName(entry.rep)).filter(Boolean)).size;

      return {
        selectedProduct,
        weeksIncluded: scorecardWeeks.length,
        repCount,
        netVolume,
        grossVolume: soldPriceAmount,
        workingAmount,
        netCount,
        closePct: issueCount > 0 ? soldCount / issueCount : 0,
        demoPct: issueCount > 0 ? demoCount / issueCount : 0,
        netPct: soldCount > 0 ? netCount / soldCount : 0,
        nsli: issueCount > 0 ? netVolume / issueCount : 0,
        gsli: issueCount > 0 ? soldPriceAmount / issueCount : 0,
      };
    }

    const scoreRows = scorecardFilteredReps.map((rep) => {
      const selectedVolume = selectedProduct === "All Products"
        ? safeNum(rep.netVolume)
        : safeNum(rep.netVolume) > 0
          ? productSelectionVolume(rep.products, selectedProduct)
          : 0;
      return {
        ...rep,
        selectedVolume,
      };
    }).filter((rep) => rep.selectedVolume > 0 || selectedProduct === "All Products");

    const totalVolume = scoreRows.reduce((sum, rep) => sum + safeNum(rep.selectedVolume), 0);
    const weightedAverage = (field) => {
      if (!scoreRows.length) return 0;
      if (totalVolume <= 0) return average(scoreRows.map((rep) => safeNum(rep[field])), 0);
      return scoreRows.reduce((sum, rep) => sum + safeNum(rep[field]) * (safeNum(rep.selectedVolume) / totalVolume), 0);
    };

    const estimatedSoldCount = scoreRows.reduce((sum, rep) => {
      const counts = estimatedLeadCounts(rep);
      return sum + safeNum(counts.soldCount);
    }, 0);
    const estimatedIssueCount = scoreRows.reduce((sum, rep) => sum + safeNum(estimatedLeadCounts(rep).issueCount), 0);
    const estimatedNetCount = scoreRows.reduce((sum, rep) => {
      const counts = estimatedLeadCounts(rep);
      if (selectedProduct === "All Products") return sum + safeNum(counts.netCount);
      return sum + (safeNum(rep.netVolume) > 0 ? safeNum(counts.netCount) * (safeNum(rep.selectedVolume) / safeNum(rep.netVolume)) : 0);
    }, 0);
    const workingAmount = weightedAverage("netPct") > 0 ? totalVolume / weightedAverage("netPct") : totalVolume;

    return {
      selectedProduct,
      weeksIncluded: scorecardWeeks.length,
      repCount: scoreRows.length,
      netVolume: totalVolume,
      grossVolume: workingAmount,
      workingAmount,
      netCount: estimatedNetCount,
      closePct: weightedAverage("closePct"),
      demoPct: weightedAverage("demoPct"),
      netPct: weightedAverage("netPct"),
      nsli: weightedAverage("nsli"),
      gsli: estimatedIssueCount > 0 ? workingAmount / estimatedIssueCount : 0,
    };
  }, [uploadData, scorecardDateRange, selectedRep, selectedGroup, selectedGroupMembers, selectedProduct, scorecardFilteredReps, scorecardWeeks]);
  const selectedRangeDays = useMemo(() => {
    if (dateRange.start && dateRange.end) {
      const start = new Date(`${dateRange.start}T00:00:00`);
      const end = new Date(`${dateRange.end}T00:00:00`);
      const diff = Math.round((end.getTime() - start.getTime()) / 86400000) + 1;
      return Math.max(1, diff);
    }
    return Math.max(1, safeNum(performanceTimeframe) || 90);
  }, [dateRange, performanceTimeframe]);

  const annualizedNetVolume = useMemo(() => {
    return safeNum(dashboardMetrics.netVolume) * (365 / Math.max(1, selectedRangeDays));
  }, [dashboardMetrics.netVolume, selectedRangeDays]);

  const goalManagementCards = useMemo(() => {
    return [
      {
        key: "closePct",
        title: "Close %",
        subtitle: "Minimum close rate target",
        actual: safeNum(dashboardMetrics.closePct),
        goal: safeNum(goalTargets.closePct),
        step: 0.01,
        inputType: "percent",
        formatter: (value) => pct(value, 0),
      },
      {
        key: "netPct",
        title: "Net %",
        subtitle: "Minimum net rate target",
        actual: safeNum(dashboardMetrics.netPct),
        goal: safeNum(goalTargets.netPct),
        step: 0.01,
        inputType: "percent",
        formatter: (value) => pct(value, 0),
      },
      {
        key: "nsli",
        title: "NSLI",
        subtitle: "Minimum net sold per issued lead",
        actual: safeNum(dashboardMetrics.nsli),
        goal: safeNum(goalTargets.nsli),
        step: 100,
        inputType: "currency",
        formatter: currency,
      },
      {
        key: "demoPct",
        title: "Demo %",
        subtitle: "Minimum demo conversion rate",
        actual: safeNum(dashboardMetrics.demoPct),
        goal: safeNum(goalTargets.demoPct),
        step: 0.01,
        inputType: "percent",
        formatter: (value) => pct(value, 0),
      },
      {
        key: "avgTicket",
        title: "Avg Ticket",
        subtitle: "Minimum average ticket",
        actual: safeNum(dashboardMetrics.avgTicket),
        goal: safeNum(goalTargets.avgTicket),
        step: 500,
        inputType: "currency",
        formatter: currency,
      },
    ];
  }, [dashboardMetrics, goalTargets]);

  const tierAnchorEnd = dateRange.end || datasetMaxDate;
  const tier90Range = useMemo(() => buildLookbackRange(90, tierAnchorEnd, datasetMinDate, datasetMaxDate), [tierAnchorEnd, datasetMinDate, datasetMaxDate]);
  const tier30Range = useMemo(() => buildLookbackRange(30, tierAnchorEnd, datasetMinDate, datasetMaxDate), [tierAnchorEnd, datasetMinDate, datasetMaxDate]);

  const tier90Dataset = useMemo(() => {
    if (uploadData && uploadData.rawEntries && uploadData.rawEntries.length) {
      return aggregateRawEntriesToDataset(uploadData.rawEntries, tier90Range.start, tier90Range.end);
    }
    const baseWeeklyDataset = uploadData?.timeframeSets?.["90"] || (uploadData?.weekLabels?.length ? uploadData : { reps: demoReps, weekLabels: demoWeeks, minDate: embeddedDemoBounds.minDate, maxDate: embeddedDemoBounds.maxDate });
    return buildDateRangedDatasetFromWeekly(baseWeeklyDataset, tier90Range.start, tier90Range.end);
  }, [uploadData, tier90Range]);

  const tier30Dataset = useMemo(() => {
    if (uploadData && uploadData.rawEntries && uploadData.rawEntries.length) {
      return aggregateRawEntriesToDataset(uploadData.rawEntries, tier30Range.start, tier30Range.end);
    }
    const baseWeeklyDataset = uploadData?.timeframeSets?.["90"] || (uploadData?.weekLabels?.length ? uploadData : { reps: demoReps, weekLabels: demoWeeks, minDate: embeddedDemoBounds.minDate, maxDate: embeddedDemoBounds.maxDate });
    return buildDateRangedDatasetFromWeekly(baseWeeklyDataset, tier30Range.start, tier30Range.end);
  }, [uploadData, tier30Range]);

  const tier90Reps = useMemo(() => dedupeAndNormalizeReps(buildMergedData(tier90Dataset), tier90Dataset?.weekLabels?.length || 16), [tier90Dataset]);
  const tier30Reps = useMemo(() => dedupeAndNormalizeReps(buildMergedData(tier30Dataset), tier30Dataset?.weekLabels?.length || 16), [tier30Dataset]);

  const tierManagementData = useMemo(() => {
    const allNames = Array.from(new Set([
      ...tier90Reps.map((rep) => canonicalName(rep.rep)),
      ...tier30Reps.map((rep) => canonicalName(rep.rep)),
      ...reps.map((rep) => canonicalName(rep.rep)),
    ].filter(Boolean)));

    const by90 = Object.fromEntries(tier90Reps.map((rep) => [canonicalName(rep.rep), rep]));
    const by30 = Object.fromEntries(tier30Reps.map((rep) => [canonicalName(rep.rep), rep]));
    const byCurrent = Object.fromEntries(reps.map((rep) => [canonicalName(rep.rep), rep]));

    const combinedReps = allNames.map((name) => {
      const rep90 = by90[name];
      const rep30 = by30[name];
      const currentRep = byCurrent[name] || rep30 || rep90;
      const ninetyTier = classifyCloseTier(rep90?.closePct ?? currentRep?.closePct, tierThresholds);
      const thirtyTier = classifyCloseTier(rep30?.closePct ?? currentRep?.closePct, tierThresholds);
      return {
        rep: currentRep?.rep || name,
        team: currentRep?.team || rep30?.team || rep90?.team || "Unknown",
        tier90: ninetyTier,
        tier30: thirtyTier,
        bucketId: `${ninetyTier}${thirtyTier}`,
        netVolume: safeNum(currentRep?.netVolume),
      };
    });

    const distribution = {
      A: combinedReps.filter((rep) => rep.tier30 === "A").length,
      B: combinedReps.filter((rep) => rep.tier30 === "B").length,
      C: combinedReps.filter((rep) => rep.tier30 === "C").length,
      total: combinedReps.length,
    };

    const bucketOrder = ["AA", "AB", "AC", "BA", "BB", "BC", "CA", "CB", "CC"];
    const buckets = bucketOrder.map((bucketId) => ({
      bucketId,
      members: combinedReps.filter((rep) => rep.bucketId === bucketId),
      count: combinedReps.filter((rep) => rep.bucketId === bucketId).length,
      ...movementBucketMeta(bucketId),
      config: movementConfig[bucketId] || DEFAULT_MOVEMENT_CONFIG[bucketId],
    }));

    return { combinedReps, distribution, buckets };
  }, [tier90Reps, tier30Reps, reps, tierThresholds, movementConfig]);

  const studioAnalyticsContext = useMemo(() => {
    const sortedByVolume = [...windowFilteredReps].sort((a, b) => safeNum(b.netVolume) - safeNum(a.netVolume));
    const focusRep = selectedRepRecord || sortedByVolume[0] || null;
    const leader = sortedByVolume[0] || null;
    const trailer = sortedByVolume.length ? sortedByVolume[sortedByVolume.length - 1] : null;
    const primaryCoaching = focusRep ? coachingForRep(focusRep, windowFilteredReps, performanceTimeframe) : [];
    const volumeValue = volumeMetric === "gross" ? safeNum(dashboardMetrics.grossVolume) : safeNum(dashboardMetrics.netVolume);
    const volumeLabel = volumeMetric === "gross" ? "gross volume" : "net volume";
    const focusLabel = selectedRep !== "All Reps"
      ? selectedRep
      : selectedGroup !== "All Groups"
        ? `${selectedGroup} team`
        : "all filtered reps";
    const topProduct = productMixData.length ? productMixData[0].name : selectedProduct;
    const weakestProduct = productMixData.length ? productMixData[productMixData.length - 1].name : selectedProduct;

    return {
      focusRep,
      leader,
      trailer,
      primaryCoaching,
      volumeValue,
      volumeLabel,
      focusLabel,
      topProduct,
      weakestProduct,
    };
  }, [windowFilteredReps, selectedRepRecord, performanceTimeframe, volumeMetric, dashboardMetrics, selectedRep, selectedGroup, productMixData, selectedProduct]);

  const buildAnalystResponse = (question = "") => {
    const prompt = cleanText(question).toLowerCase();
    const focusRep = studioAnalyticsContext.focusRep;
    const coaching = studioAnalyticsContext.primaryCoaching || [];
    const lines = [];

    lines.push(`Scope: ${studioAnalyticsContext.focusLabel} • ${selectedProduct} • ${effectiveRangeLabel}.`);
    lines.push(`Current ${studioAnalyticsContext.volumeLabel}: ${currency(studioAnalyticsContext.volumeValue)}. Close rate: ${pct(dashboardMetrics.closePct, 0)}. Net: ${pct(dashboardMetrics.netPct)}. NSLI: ${currency(dashboardMetrics.nsli)}. Demo: ${pct(dashboardMetrics.demoPct)}.`);

    if (studioAnalyticsContext.topProduct && studioAnalyticsContext.topProduct !== "All Products") {
      lines.push(`Best-performing product in the current view: ${studioAnalyticsContext.topProduct}. Weakest product exposure in the current ranking: ${studioAnalyticsContext.weakestProduct}.`);
    }

    if (focusRep) {
      lines.push(`Primary rep focus: ${focusRep.rep} (${focusRep.team}). Movement profile: ${focusRep.movement}. Avg ticket: ${currency(avgTicketFromRep(focusRep))}.`);
    }

    const priorities = [];
    if (dashboardMetrics.closePct < goalTargets.closePct) priorities.push("Raise close rate by tightening discovery, urgency, and the ask.");
    if (dashboardMetrics.demoPct < goalTargets.demoPct) priorities.push("Improve demo quality and set control so more issued leads become real selling opportunities.");
    if (dashboardMetrics.netPct < goalTargets.netPct) priorities.push("Reduce fallout with cleaner expectation setting, financing alignment, and better next-step control.");
    if (dashboardMetrics.nsli < goalTargets.nsli) priorities.push("Lift NSLI with stronger premium anchoring, better package framing, and tighter product-fit recommendations.");
    if (!priorities.length) priorities.push("Core KPI mix is healthy. Focus on replicating what is already working across more reps and product categories.");

    lines.push(`Priority actions: ${priorities.slice(0, 2).join(" ")}`);

    if (coaching.length) {
      const matched = coaching.find((item) => item.framework === "Gap Selling") || coaching[0];
      if (matched) {
        lines.push(`${matched.framework}: ${matched.guidance}`);
      }
    }

    if (prompt.includes("compare") && studioAnalyticsContext.leader && studioAnalyticsContext.trailer) {
      lines.push(`Comparison: ${studioAnalyticsContext.leader.rep} leads the current view at ${currency(studioAnalyticsContext.leader.netVolume)} net volume, while ${studioAnalyticsContext.trailer.rep} trails at ${currency(studioAnalyticsContext.trailer.netVolume)}.`);
    }

    if (prompt.includes("team") && selectedGroup !== "All Groups") {
      lines.push(`Team read: ${selectedGroup} currently has ${windowFilteredReps.length} reps in scope. Use the refresh button after changing group or rep filters to regenerate this coaching summary against the current view.`);
    }

    return lines.join(String.fromCharCode(10, 10));
  };

  const refreshAnalystRecommendations = () => {
    setAnalystMessages([
      {
        role: "assistant",
        content: buildAnalystResponse(),
      },
    ]);
  };

  const productTabs = ["All Products", "Baths", "Windows", "Doors", "Roofing", "Siding", "Roofing + Siding", "Windows + Doors"];

  async function onUpload(event) {
    const file = event.target.files && event.target.files[0];
    if (!file) return;

    const arrayBuffer = await file.arrayBuffer();
    const parsed = parseUploadedArrayBuffer(arrayBuffer);

    const parsedMinDate = parsed?.minDate || embeddedDemoBounds.minDate;
    const parsedMaxDate = parsed?.maxDate || embeddedDemoBounds.maxDate || toIsoDate(new Date());
    const initialRange = { start: parsedMinDate, end: parsedMaxDate };

    setUploadMeta({ workbookName: file.name, sheetNames: parsed.sheetNames || [] });
    setUploadData(parsed);
    setSelectedRep("All Reps");
    setRepSearch("");
    setDateRange(initialRange);
    setDashboardDraftRange(initialRange);
    setScorecardDateRange(initialRange);
    setScorecardDraftRange(initialRange);
    setProjectorManualRange(initialRange);
    setProjectorDraftRange(initialRange);
    setDashboardRangeEditorOpen(false);
    setScorecardRangeEditorOpen(false);
    setProjectorRangeEditorOpen(false);
  }


  return (
    <div
      className="min-h-screen bg-[var(--app-bg)] text-[var(--app-text)] tracking-[-0.01em]"
      style={{
        ...themeVars,
        fontFamily: "'Rajdhani', 'Inter', ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif",
      }}
    >
      <div className="grid min-h-screen grid-cols-12 lg:grid-cols-[auto_minmax(0,1fr)]">
        <aside className={`col-span-12 border-b border-[var(--sidebar-border)] bg-[var(--sidebar-bg)] p-4 transition-all duration-200 lg:sticky lg:top-0 lg:col-span-1 lg:h-screen lg:self-start lg:overflow-y-auto lg:border-b-0 lg:border-r lg:pb-10 ${isSidebarCollapsed ? "lg:w-[88px]" : "lg:w-[280px]"}`}>
          <div className="flex items-start justify-between gap-2">
            <div className="flex min-w-0 flex-1 items-center gap-2">
              <div className="shrink-0 rounded-2xl border border-lime-500/40 bg-lime-500/10 px-2.5 py-2 text-[var(--kpi-good)]">
                <Target className="h-4.5 w-4.5" />
              </div>
              <div className="min-w-0 flex-1">
                {!isSidebarCollapsed ? (
                  <>
                    <div className="w-full text-[1.55rem] font-black uppercase leading-[0.9] tracking-[0.06em] text-[#7ed14b]" style={{ fontFamily: "'Rajdhani', 'Inter', sans-serif" }}>
                      OVERWATCH
                    </div>
                    <div className="mt-1">
                      <Select value={activeDepartment} onValueChange={setActiveDepartment}>
                        <SelectTrigger className="h-8 border-[var(--border)] bg-[var(--panel-bg)] text-[12px] tracking-[0.08em] text-[var(--text-soft)]" style={{ fontFamily: "'Rajdhani', 'Inter', sans-serif" }}>
                          <SelectValue />
                        </SelectTrigger>
                        <SelectContent>
                          {[
                            "Sales Department",
                            "Marketing",
                            "Finance",
                            "Production",
                            "Admin",
                            "Human Resources",
                          ].map((department) => (
                            <SelectItem key={department} value={department}>{department}</SelectItem>
                          ))}
                        </SelectContent>
                      </Select>
                    </div>
                  </>
                ) : (
                  <div className="w-full text-center text-[0.95rem] font-black uppercase leading-[0.9] tracking-[0.08em] text-[#7ed14b]" style={{ fontFamily: "'Rajdhani', 'Inter', sans-serif" }}>
                    OW
                  </div>
                )}
              </div>
            </div>
            <Button
              type="button"
              variant="outline"
              onClick={() => setIsSidebarCollapsed((current) => !current)}
              className="h-9 w-9 shrink-0 border-[var(--border)] bg-[var(--panel-bg)] px-0 text-[var(--text-soft)] hover:bg-[var(--button-hover)]"
            >
              <span className="text-sm font-bold">{isSidebarCollapsed ? "→" : "←"}</span>
            </Button>
          </div>

          {!isSidebarCollapsed ? (
            activeDepartment === "Sales Department" ? (
              <>
                <div className="mt-4 rounded-2xl border border-[var(--border-strong)] bg-[var(--sidebar-card-bg)]/96 p-4">
                  <div className="text-sm font-semibold uppercase tracking-[0.25em] text-[#9aa7bb]">Workspace</div>
                  <div className="mt-4 space-y-2">
                    <button
                      type="button"
                      onClick={() => setActiveWorkspace("reviewStudio")}
                      className={`flex w-full items-center justify-between rounded-xl border px-3 py-2 text-left text-sm font-semibold ${activeWorkspace === "reviewStudio" ? "border-[var(--toggle-active-border)] bg-[var(--toggle-active-bg)] text-[var(--toggle-active-text)]" : "border-[var(--border)] bg-[var(--panel-bg)] text-[var(--text-soft)] hover:bg-[var(--button-hover)]"}`}
                    >
                      <span>Review Studio</span>
                      <span className="text-[10px] uppercase tracking-[0.18em]">Live</span>
                    </button>
                  </div>
                </div>

                <div className="mt-4 rounded-2xl border border-[var(--border-strong)] bg-[var(--sidebar-card-bg)]/96 p-4">
                  <button
                    type="button"
                    onClick={() => setShowFiltersPanel((current) => !current)}
                    className="mb-0 flex w-full items-center justify-between text-left"
                  >
                    <div className="text-sm font-semibold uppercase tracking-[0.25em] text-[#9aa7bb]">Filters</div>
                    <div className="rounded-lg border border-[var(--border)] bg-[var(--panel-bg)] px-2 py-1 text-xs text-[var(--text-soft)]">
                      {showFiltersPanel ? "Hide" : "Show"}
                    </div>
                  </button>
                  {showFiltersPanel ? (
                    <div className="mt-4 space-y-4">
                      <div>
                        <Label className="mb-2 block">Group</Label>
                        <div className="flex items-center gap-2">
                          <Select
                            value={selectedGroup}
                            onValueChange={(value) => {
                              setSelectedGroup(value);
                              setSelectedRep("All Reps");
                              if (value === "All Groups") {
                                setGroupNameInput("");
                                setGroupDraftMembers([]);
                              } else {
                                setGroupNameInput(value);
                                setGroupDraftMembers(effectiveGroups[value] || []);
                              }
                            }}
                          >
                            <SelectTrigger className="border-[var(--border)] bg-[var(--panel-bg)]"><SelectValue /></SelectTrigger>
                            <SelectContent>{groupOptions.map((group) => <SelectItem key={group.value} value={group.value}>{group.label}</SelectItem>)}</SelectContent>
                          </Select>
                          <Button
                            type="button"
                            variant="outline"
                            disabled={selectedGroup === "All Groups" || PRELOADED_GROUP_NAMES.includes(selectedGroup)}
                            onClick={() => {
                              if (selectedGroup === "All Groups" || PRELOADED_GROUP_NAMES.includes(selectedGroup)) return;
                              setCustomGroups((current) => {
                                const next = { ...current };
                                delete next[selectedGroup];
                                return next;
                              });
                              setSelectedGroup("All Groups");
                              setSelectedRep("All Reps");
                              setGroupNameInput("");
                              setGroupDraftMembers([]);
                            }}
                            className="border-[var(--border)] bg-[var(--panel-bg)] text-[var(--text-soft)] hover:bg-slate-800"
                          >
                            Delete
                          </Button>
                        </div>
                      </div>
                      <div>
                        <Label className="mb-2 block">Rep</Label>
                        <Input
                          value={repSearch}
                          onChange={(event) => setRepSearch(event.target.value)}
                          placeholder="Search reps"
                          className="mb-2 border-[var(--border)] bg-[var(--panel-bg)] text-[var(--text-strong)] placeholder:text-[var(--kpi-title)]"
                        />
                        <div className="rounded-xl border border-[var(--border)] bg-[var(--panel-bg)]/70">
                          <button
                            type="button"
                            onClick={() => setSelectedRep("All Reps")}
                            className={`w-full border-b border-[var(--border)] px-3 py-2 text-left text-sm ${selectedRep === "All Reps" ? "bg-[var(--selection-bg)] text-[var(--selection-text)]" : "text-[var(--text-soft)] hover:bg-slate-800"}`}
                          >
                            All Reps
                          </button>
                          <ScrollArea className="h-48">
                            <div className="p-1">
                              {searchedRepOptions.filter((rep) => rep.value !== "All Reps").map((rep) => (
                                <button
                                  key={rep.value}
                                  type="button"
                                  onClick={() => setSelectedRep(rep.value)}
                                  className={`mb-1 w-full rounded-lg px-3 py-2 text-left text-sm ${selectedRep === rep.value ? "bg-[var(--selection-bg)] text-[var(--selection-text)]" : "text-[var(--text-soft)] hover:bg-slate-800"}`}
                                >
                                  {rep.label}
                                </button>
                              ))}
                              {searchedRepOptions.filter((rep) => rep.value !== "All Reps").length === 0 ? (
                                <div className="px-3 py-2 text-sm text-[var(--kpi-title)]">No matching reps</div>
                              ) : null}
                            </div>
                          </ScrollArea>
                        </div>
                      </div>
                    </div>
                  ) : null}
                </div>

                <div className="mt-4 rounded-2xl border border-[var(--border-strong)] bg-[var(--sidebar-card-bg)]/96 p-4 lg:mb-6">
                  <div className="text-sm font-semibold uppercase tracking-[0.25em] text-[#9aa7bb]">Settings</div>
                  <div className="mt-4 space-y-2">
                    <button
                      type="button"
                      onClick={() => setActiveWorkspace("goalManagement")}
                      className={`flex w-full items-center justify-between rounded-xl border px-3 py-2 text-left text-sm font-semibold ${activeWorkspace === "goalManagement" ? "border-[var(--toggle-active-border)] bg-[var(--toggle-active-bg)] text-[var(--toggle-active-text)]" : "border-[var(--border)] bg-[var(--panel-bg)] text-[var(--text-soft)] hover:bg-[var(--button-hover)]"}`}
                    >
                      <span>Goal Management</span>
                      <span className="text-[10px] uppercase tracking-[0.18em]">Targets</span>
                    </button>
                    <button
                      type="button"
                      onClick={() => setActiveWorkspace("tierManagement")}
                      className={`flex w-full items-center justify-between rounded-xl border px-3 py-2 text-left text-sm font-semibold ${activeWorkspace === "tierManagement" ? "border-[var(--toggle-active-border)] bg-[var(--toggle-active-bg)] text-[var(--toggle-active-text)]" : "border-[var(--border)] bg-[var(--panel-bg)] text-[var(--text-soft)] hover:bg-[var(--button-hover)]"}`}
                    >
                      <span>Tier Management</span>
                      <span className="text-[10px] uppercase tracking-[0.18em]">Movement</span>
                    </button>
                    <button
                      type="button"
                      onClick={() => setActiveWorkspace("teamManagement")}
                      className={`flex w-full items-center justify-between rounded-xl border px-3 py-2 text-left text-sm font-semibold ${activeWorkspace === "teamManagement" ? "border-[var(--toggle-active-border)] bg-[var(--toggle-active-bg)] text-[var(--toggle-active-text)]" : "border-[var(--border)] bg-[var(--panel-bg)] text-[var(--text-soft)] hover:bg-[var(--button-hover)]"}`}
                    >
                      <span>Team Management</span>
                      <span className="text-[10px] uppercase tracking-[0.18em]">Assign</span>
                    </button>
                  </div>
                </div>
              </>
            ) : (
              <div className="mt-4 rounded-2xl border border-[var(--border-strong)] bg-[var(--sidebar-card-bg)]/96 p-4 text-sm text-[var(--text-soft)]">
                <div className="text-sm font-semibold uppercase tracking-[0.25em] text-[#9aa7bb]">Department Status</div>
                <div className="mt-3 rounded-xl border border-[var(--border)] bg-[var(--panel-bg)] p-3">
                  <div className="font-semibold text-[var(--text-strong)]">{activeDepartment}</div>
                  <div className="mt-1 text-xs text-[var(--kpi-title)]">Under construction</div>
                </div>
              </div>
            )
          ) : (
            <div className="mt-4 rounded-2xl border border-[var(--border-strong)] bg-[var(--sidebar-card-bg)]/96 p-4 text-center text-xs text-[var(--kpi-title)]">
              <div className="font-semibold text-[var(--text-strong)]">{activeDepartment === "Sales Department" ? "Sales" : activeDepartment}</div>
              <div className="mt-1">Collapsed</div>
            </div>
          )}
        </aside>

        <main className="col-span-12 px-4 pb-4 pt-1 lg:col-span-1 lg:px-6 lg:pb-6 lg:pt-1.5">
          <div className={activeDepartment !== "Sales Department" ? "block" : "hidden"}>
            <div className="sticky top-0 z-30 -mx-4 px-4 pb-1.5 backdrop-blur lg:-mx-6 lg:px-6" style={{ backgroundColor: "var(--header-bg)" }}>
              <div>
                <h1 className="text-lg font-semibold tracking-[-0.01em] text-[var(--text-strong)]" style={{ fontFamily: "'Rajdhani', 'Inter', sans-serif" }}>{activeDepartment}</h1>
                <p className="text-[11px] text-[#9ba8bb]">This department view is currently under construction.</p>
              </div>
            </div>

            <div className="mt-8 max-w-3xl">
              <Card className="border-[var(--border-strong)] bg-[var(--card-bg)]">
                <CardContent className="p-8">
                  <div className="text-xs font-semibold uppercase tracking-[0.25em] text-[#9aa7bb]">Module Status</div>
                  <div className="mt-4 text-3xl font-bold text-[var(--text-strong)]">{activeDepartment}</div>
                  <div className="mt-2 text-base text-[var(--text-soft)]">This department module is under construction. Sales Department remains the active, fully functional workspace.</div>
                  <div className="mt-6 grid grid-cols-1 gap-3 sm:grid-cols-2">
                    {[
                      "Department KPI scorecards",
                      "Workflow-specific analytics",
                      "Team and goal controls",
                      "Cross-functional review tools",
                    ].map((item) => (
                      <div key={item} className="rounded-xl border border-[var(--border)] bg-[var(--panel-bg)] px-4 py-3 text-sm text-[var(--text-soft)]">
                        {item}
                        <div className="mt-1 text-xs text-[var(--kpi-title)]">Under construction</div>
                      </div>
                    ))}
                  </div>
                </CardContent>
              </Card>
            </div>
          </div>

          <div className={activeDepartment === "Sales Department" && activeWorkspace === "reviewStudio" ? "block" : "hidden"}>
              <div className="sticky top-0 z-30 -mx-4 px-4 pb-1.5 backdrop-blur lg:-mx-6 lg:px-6" style={{ backgroundColor: "var(--header-bg)" }}>
                <div className="flex flex-col gap-1.5 xl:flex-row xl:items-start xl:justify-between">
                <div>
                  <h1 className="text-lg font-semibold tracking-[-0.01em] text-[var(--text-strong)]" style={{ fontFamily: "'Rajdhani', 'Inter', sans-serif" }}>Overwatch Analytics</h1>
                  <p className="text-[11px] text-[#9ba8bb]">Interactive KPI review, coaching recommendations, projections, and live insights across any imported date range.</p>
                </div>
                <div className="flex w-full max-w-[520px] flex-col items-stretch gap-1 xl:items-end">
                  <div className="flex flex-wrap items-center justify-end gap-2">
                    <a
                      href="https://improveit360-9618.my.salesforce.com/00OPf000007Ws9J"
                      target="_blank"
                      rel="noreferrer"
                      className="inline-flex items-center justify-center px-4 text-[13px] font-semibold transition hover:bg-white"
                      style={actionPillStyle}
                    >
                      Data Source
                    </a>
                    <label
                      className="inline-flex cursor-pointer items-center justify-center gap-2 px-4 text-[13px] font-semibold transition hover:bg-white"
                      style={actionPillStyle}
                    >
                      <UploadIcon className="h-4 w-4" />
                      <span>Upload New File</span>
                      <input type="file" accept=".xlsx,.xls" className="hidden" onChange={onUpload} />
                    </label>
                    <button
                      type="button"
                      onClick={() => {
                        if (typeof document !== "undefined") {
                          document.getElementById("period-kpi-scorecard")?.scrollIntoView({ behavior: "smooth", block: "start" });
                        }
                      }}
                      className="inline-flex items-center justify-center gap-2 px-4 text-[13px] font-semibold transition hover:bg-white"
                      style={actionPillStyle}
                    >
                      <InfoIcon className="h-4 w-4" />
                      <span>KPI Info</span>
                    </button>
                    <button
                      type="button"
                      onClick={() => setIsDarkMode((current) => !current)}
                      className="inline-flex items-center justify-center gap-2 px-4 text-[13px] font-semibold transition hover:bg-white"
                      style={actionPillStyle}
                    >
                      <DarkModeIcon className="h-4 w-4" />
                      <span>{isDarkMode ? "Light Mode" : "Dark Mode"}</span>
                    </button>
                  </div>
                  <div className="text-right text-[10px] text-[var(--kpi-title)]">Loaded: {uploadMeta.workbookName}</div>
                </div>
              </div>

              <div className="mt-2 flex flex-col gap-1.5 xl:flex-row xl:items-center xl:justify-between">
                <div className="flex flex-wrap items-center gap-3">
                  {["7", "30", "60", "90"].map((option) => {
                    const isActive = performanceTimeframe === option && !dashboardRangeEditorOpen;
                    return (
                      <button
                        key={option}
                        type="button"
                        onClick={() => {
                          const nextRange = buildLookbackRange(option, dateRange.end || datasetMaxDate, datasetMinDate, datasetMaxDate);
                          setPerformanceTimeframe(option);
                          setDateRange(nextRange);
                          setDashboardDraftRange(nextRange);
                          setDashboardRangeEditorOpen(false);
                        }}
                        className={`min-w-[62px] rounded-xl border px-3 py-1.5 text-[13px] font-semibold ${isActive ? "border-[var(--toggle-active-border)] bg-[var(--toggle-active-bg)] text-[var(--toggle-active-text)]" : "border-[var(--border)] bg-[var(--button-bg)] text-[#a9b4c5] hover:bg-[var(--button-hover)]"}`}
                      >
                        {option}D
                      </button>
                    );
                  })}
                  <button
                    type="button"
                    onClick={() => {
                      setDashboardDraftRange({
                        start: dateRange.start || datasetMinDate,
                        end: dateRange.end || datasetMaxDate,
                      });
                      setDashboardRangeEditorOpen(true);
                    }}
                    className={`inline-flex items-center justify-center gap-2 rounded-xl border px-4 py-2 text-base font-semibold ${dashboardRangeEditorOpen || performanceTimeframe === "manual" ? "border-[#7a8ea9] bg-[var(--button-active-bg)] text-[var(--text-strong)] ring-1 ring-[var(--toggle-active-ring)]/70" : "border-[var(--border)] bg-[var(--panel-bg)] text-[var(--text-soft)] hover:bg-[var(--button-hover)]"}`}
                  >
                    <CalendarDays className="h-4 w-4" />
                    {effectiveRangeLabel}
                  </button>
                </div>
                <div className="flex flex-wrap items-center gap-2 xl:justify-end">
                  <div className="mr-2 flex items-center gap-2 rounded-xl border border-[var(--border)] bg-[var(--panel-bg)] p-1">
                    <button
                      type="button"
                      onClick={() => setVolumeMetric("net")}
                      className={`rounded-lg px-3 py-1.5 text-xs font-semibold ${volumeMetric === "net" ? "bg-[var(--button-active-bg)] text-[var(--text-strong)]" : "text-[#a9b4c5] hover:bg-[var(--button-bg)]"}`}
                    >
                      Net
                    </button>
                    <button
                      type="button"
                      onClick={() => setVolumeMetric("gross")}
                      className={`rounded-lg px-3 py-1.5 text-xs font-semibold ${volumeMetric === "gross" ? "bg-[var(--button-active-bg)] text-[var(--text-strong)]" : "text-[#a9b4c5] hover:bg-[var(--button-bg)]"}`}
                    >
                      Gross
                    </button>
                  </div>
                  {productTabs.map((product) => (
                    <button
                      key={product}
                      type="button"
                      onClick={() => setSelectedProduct(product)}
                      className={`rounded-xl border px-3 py-1.5 text-xs font-semibold ${selectedProduct === product ? "border-[#7a8ea9] bg-[var(--button-active-bg)] text-[var(--text-strong)] ring-1 ring-[var(--toggle-active-ring)]/55" : "border-[var(--border)] bg-[var(--button-bg)] text-[#a9b4c5] hover:bg-[var(--button-hover)]"}`}
                    >
                      {product}
                    </button>
                  ))}
                </div>
              </div>

              {dashboardRangeEditorOpen ? (
                <div className="mt-4 max-w-[560px] rounded-[28px] border border-[var(--border-strong)] bg-[var(--card-bg)]/95 p-5 shadow-2xl shadow-slate-950/40">
                  <div className="text-xs font-semibold uppercase tracking-[0.25em] text-[#9aa7bb]">Custom Date Range</div>
                  <div className="mt-4 space-y-4">
                    <div>
                      <Label className="mb-2 block text-base text-[var(--text-soft)]">From</Label>
                      <DatePickerField
                        value={dashboardDraftRange.start}
                        onChange={(value) => setDashboardDraftRange((current) => ({ ...current, start: value }))}
                        maxDate={dashboardDraftRange.end || datasetMaxDate}
                        minDate={datasetMinDate}
                        placeholder="Choose start date"
                      />
                    </div>
                    <div>
                      <Label className="mb-2 block text-base text-[var(--text-soft)]">To</Label>
                      <DatePickerField
                        value={dashboardDraftRange.end}
                        onChange={(value) => setDashboardDraftRange((current) => ({ ...current, end: value }))}
                        minDate={dashboardDraftRange.start || datasetMinDate}
                        maxDate={datasetMaxDate}
                        placeholder="Choose end date"
                      />
                    </div>
                    <div className="flex items-center gap-3">
                      <Button
                        type="button"
                        onClick={() => {
                          const appliedRange = {
                            start: clampIsoDate(dashboardDraftRange.start || datasetMinDate, datasetMinDate, datasetMaxDate),
                            end: clampIsoDate(dashboardDraftRange.end || datasetMaxDate, datasetMinDate, datasetMaxDate),
                          };
                          setPerformanceTimeframe("manual");
                          setDateRange(appliedRange);
                          setDashboardDraftRange(appliedRange);
                          setDashboardRangeEditorOpen(false);
                        }}
                        className="min-w-[160px] rounded-2xl bg-lime-600 text-lg font-semibold text-slate-950 hover:bg-lime-500"
                      >
                        Apply
                      </Button>
                      <Button
                        type="button"
                        variant="outline"
                        onClick={() => {
                          setDashboardDraftRange({
                            start: dateRange.start || datasetMinDate,
                            end: dateRange.end || datasetMaxDate,
                          });
                          setDashboardRangeEditorOpen(false);
                        }}
                        className="rounded-2xl border-[var(--border)] bg-[var(--panel-bg)]/60 text-lg text-[var(--text-soft)] hover:bg-slate-800"
                      >
                        Cancel
                      </Button>
                    </div>
                  </div>
                </div>
              ) : null}

              <div className="mt-4 grid grid-cols-1 gap-2 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-6">
                <StatCard
                  title={volumeMetric === "gross" ? "Gross Volume" : "Net Volume"}
                  value={currency(volumeMetric === "gross" ? stickyHeaderMetrics.ytd.grossVolume : stickyHeaderMetrics.ytd.netVolume)}
                  subvalue="YTD"
                  secondaryValue={`MTD ${currency(volumeMetric === "gross" ? stickyHeaderMetrics.mtd.grossVolume : stickyHeaderMetrics.mtd.netVolume)}`}
                  secondaryRawValue={null}
                  tertiaryValue={volumeMetric === "net" ? `Annual pace ${currency(stickyYtdPace.annualized)}` : ""}
                  tertiaryClassName="text-[var(--kpi-goal)]"
                  icon={DollarSign}
                  accent={volumeMetric === "gross" ? "text-[var(--text-strong)]" : "text-[var(--kpi-volume-primary)]"}
                  rawValue={null}
                />
                <StatCard title="Close %" value={pct(stickyHeaderMetrics.ytd.closePct, 0)} subvalue="YTD" secondaryValue={`MTD ${pct(stickyHeaderMetrics.mtd.closePct, 0)}`}
                  secondaryRawValue={stickyHeaderMetrics.mtd.closePct} icon={Target} rawValue={stickyHeaderMetrics.ytd.closePct} goalNote={`Goal ${pct(goalTargets.closePct, 0)}`} goals={goalTargets} bands={kpiColorBands} />
                <StatCard title="Net %" value={pct(stickyHeaderMetrics.ytd.netPct)} subvalue="YTD" secondaryValue={`MTD ${pct(stickyHeaderMetrics.mtd.netPct)}`}
                  secondaryRawValue={stickyHeaderMetrics.mtd.netPct} icon={TrendingUp} rawValue={stickyHeaderMetrics.ytd.netPct} goalNote={`Goal ${pct(goalTargets.netPct)}`} goals={goalTargets} bands={kpiColorBands} />
                <StatCard title="NSLI" value={currency(stickyHeaderMetrics.ytd.nsli)} subvalue="YTD" secondaryValue={`MTD ${currency(stickyHeaderMetrics.mtd.nsli)}`}
                  secondaryRawValue={stickyHeaderMetrics.mtd.nsli} icon={Trophy} rawValue={stickyHeaderMetrics.ytd.nsli} goalNote={`Goal ${currency(goalTargets.nsli)}`} goals={goalTargets} bands={kpiColorBands} />
                <StatCard title="Demo %" value={pct(stickyHeaderMetrics.ytd.demoPct)} subvalue="YTD" secondaryValue={`MTD ${pct(stickyHeaderMetrics.mtd.demoPct)}`}
                  secondaryRawValue={stickyHeaderMetrics.mtd.demoPct} icon={TrendingUp} rawValue={stickyHeaderMetrics.ytd.demoPct} goalNote={`Goal ${pct(goalTargets.demoPct)}`} goals={goalTargets} bands={kpiColorBands} />
                <StatCard title="Avg Ticket" value={currency(stickyHeaderMetrics.ytd.avgTicket)} subvalue="YTD" secondaryValue={`MTD ${currency(stickyHeaderMetrics.mtd.avgTicket)}`}
                  secondaryRawValue={stickyHeaderMetrics.mtd.avgTicket} icon={Target} rawValue={stickyHeaderMetrics.ytd.avgTicket} />
              </div>
            </div>

              <div className="mt-6 grid grid-cols-1 gap-4 xl:grid-cols-12">
                <Card className="xl:col-span-8 border-[var(--border-strong)] bg-[var(--card-bg)]">
                  <CardHeader>
                    <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
                      <CardTitle className="text-sm uppercase tracking-[0.25em] text-[#9aa7bb]">Weekly Trend</CardTitle>
                      <div className="flex items-center gap-2">
                        <button
                          type="button"
                          onClick={() => setWeeklyTrendMetric("netSales")}
                          className={`rounded-xl border px-3 py-1.5 text-xs font-semibold ${weeklyTrendMetric === "netSales" ? "border-[var(--toggle-active-border)] bg-[var(--toggle-active-bg)] text-[var(--toggle-active-text)]" : "border-[var(--border)] bg-[var(--button-bg)] text-[#a9b4c5] hover:bg-[var(--button-hover)]"}`}
                        >
                          Net Sales
                        </button>
                        <button
                          type="button"
                          onClick={() => setWeeklyTrendMetric("volume")}
                          className={`rounded-xl border px-3 py-1.5 text-xs font-semibold ${weeklyTrendMetric === "volume" ? "border-[var(--toggle-active-border)] bg-[var(--toggle-active-bg)] text-[var(--toggle-active-text)]" : "border-[var(--border)] bg-[var(--button-bg)] text-[#a9b4c5] hover:bg-[var(--button-hover)]"}`}
                        >
                          Volume
                        </button>
                      </div>
                    </div>
                  </CardHeader>
                  <CardContent className="h-[470px]">
                    <ResponsiveContainer width="100%" height="100%">
                      <BarChart data={weeklyTrendDisplaySeries} margin={{ top: 28, right: 16, left: 6, bottom: 68 }} barGap={weeklyTrendMetric === "volume" ? 3 : 6}>
                        <CartesianGrid stroke="#1e293b" vertical={false} />
                        <XAxis
                          dataKey="label"
                          stroke="#94a3b8"
                          interval={0}
                          height={58}
                          angle={-33}
                          textAnchor="end"
                          tick={{ fill: "#94a3b8", fontSize: 12 }}
                        />
                        <YAxis stroke="#94a3b8" tick={{ fill: "#94a3b8", fontSize: 12 }} tickFormatter={weeklyTrendAxisFormatter} />
                        <Tooltip
                          contentStyle={tooltipContentStyle}
                          labelStyle={tooltipLabelStyle}
                          itemStyle={tooltipItemStyle}
                          formatter={(value, name) => {
                            if (weeklyTrendMetric === "volume") {
                              return [currency(value), name === "grossVolume" ? "Gross Volume" : "Net Volume"];
                            }
                            return [weeklyTrendTooltipFormatter(value), "Net Sales"];
                          }}
                        />
                        {weeklyTrendMetric === "volume" ? (
                          <>
                            <Bar dataKey="grossVolume" radius={[8, 8, 0, 0]} fill={isDarkMode ? "#64748b" : "#94a3b8"} barSize={20} />
                            <Bar
                              dataKey="grossVolume"
                              fill="transparent"
                              stroke="transparent"
                              legendType="none"
                              isAnimationActive={false}
                              barSize={0}
                              label={({ x, y, width, value }) => (
                                <text
                                  x={safeNum(x) + safeNum(width) / 2}
                                  y={safeNum(y) - 8}
                                  fill="var(--chart-label)"
                                  opacity={0.25}
                                  fontSize={14}
                                  fontWeight="700"
                                  textAnchor="middle"
                                >
                                  {currency(value)}
                                </text>
                              )}
                            />
                            <Bar
                              dataKey="netVolume"
                              radius={[8, 8, 0, 0]}
                              fill="#38bdf8"
                              barSize={20}
                            />
                            <Bar
                              dataKey="netVolume"
                              fill="transparent"
                              stroke="transparent"
                              legendType="none"
                              isAnimationActive={false}
                              barSize={0}
                              label={({ x, y, width, value }) => (
                                <text
                                  x={safeNum(x) + safeNum(width) / 2}
                                  y={safeNum(y) - 8}
                                  fill="var(--chart-label)"
                                  fontSize={14}
                                  fontWeight="700"
                                  textAnchor="middle"
                                >
                                  {currency(value)}
                                </text>
                              )}
                            />
                          </>
                        ) : (
                          <Bar
                            dataKey="metricValue"
                            radius={[8, 8, 0, 0]}
                            fill="#38bdf8"
                            label={({ x, y, width, value }) => (
                              <text
                                x={safeNum(x) + safeNum(width) / 2}
                                y={safeNum(y) - 8}
                                fill="var(--chart-label)"
                                fontSize={14}
                                fontWeight="700"
                                textAnchor="middle"
                              >
                                {weeklyTrendMetric === "netSales"
                                  ? `${Math.round(safeNum(value))}`
                                  : currency(value)}
                              </text>
                            )}
                          />
                        )}
                      </BarChart>
                    </ResponsiveContainer>
                  </CardContent>
                </Card>

                <Card id="period-kpi-scorecard" className="xl:col-span-4 border-[var(--border-strong)] bg-[var(--card-bg)]">
                  <CardHeader>
                    <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
                      <CardTitle className="text-sm uppercase tracking-[0.25em] text-[#9aa7bb]">Period KPI Scorecard</CardTitle>
                      <button
                        type="button"
                        onClick={() => {
                          setScorecardDraftRange({
                            start: scorecardDateRange.start || datasetMinDate,
                            end: scorecardDateRange.end || datasetMaxDate,
                          });
                          setScorecardRangeEditorOpen((open) => !open);
                        }}
                        className={`inline-flex items-center justify-center gap-2 rounded-xl border px-3 py-1.5 text-xs font-semibold ${scorecardRangeEditorOpen ? "border-[#7a8ea9] bg-[var(--button-active-bg)] text-[var(--text-strong)] ring-1 ring-[var(--toggle-active-ring)]/70" : "border-[var(--border)] bg-[var(--panel-bg)] text-[var(--text-soft)] hover:bg-[var(--button-hover)]"}`}
                      >
                        <CalendarDays className="h-3.5 w-3.5" />
                        {scorecardRangeLabel}
                      </button>
                    </div>
                  </CardHeader>
                  <CardContent className="space-y-4">
                    {scorecardRangeEditorOpen ? (
                      <div className="rounded-2xl border border-[var(--border-strong)] bg-[var(--panel-bg)]/80 p-4">
                        <div className="text-xs font-semibold uppercase tracking-[0.25em] text-[#9aa7bb]">Scorecard Date Range</div>
                        <div className="mt-4 grid grid-cols-1 gap-3 sm:grid-cols-2">
                          <div>
                            <Label className="mb-2 block text-sm text-[var(--text-soft)]">From</Label>
                            <DatePickerField
                              value={scorecardDraftRange.start}
                              onChange={(value) => setScorecardDraftRange((current) => ({ ...current, start: value }))}
                              maxDate={scorecardDraftRange.end || datasetMaxDate}
                              minDate={datasetMinDate}
                              placeholder="Choose scorecard start date"
                            />
                          </div>
                          <div>
                            <Label className="mb-2 block text-sm text-[var(--text-soft)]">To</Label>
                            <DatePickerField
                              value={scorecardDraftRange.end}
                              onChange={(value) => setScorecardDraftRange((current) => ({ ...current, end: value }))}
                              minDate={scorecardDraftRange.start || datasetMinDate}
                              maxDate={datasetMaxDate}
                              placeholder="Choose scorecard end date"
                            />
                          </div>
                        </div>
                        <div className="mt-4 flex flex-wrap items-center gap-2">
                          <Button
                            type="button"
                            onClick={() => {
                              const appliedRange = {
                                start: clampIsoDate(scorecardDraftRange.start || datasetMinDate, datasetMinDate, datasetMaxDate),
                                end: clampIsoDate(scorecardDraftRange.end || datasetMaxDate, datasetMinDate, datasetMaxDate),
                              };
                              setScorecardDateRange(appliedRange);
                              setScorecardDraftRange(appliedRange);
                              setScorecardRangeEditorOpen(false);
                            }}
                            className="rounded-xl bg-lime-600 text-slate-950 hover:bg-lime-500"
                          >
                            Apply Scorecard Range
                          </Button>
                          <Button
                            type="button"
                            variant="outline"
                            onClick={() => {
                              setScorecardDraftRange({
                                start: scorecardDateRange.start || datasetMinDate,
                                end: scorecardDateRange.end || datasetMaxDate,
                              });
                              setScorecardRangeEditorOpen(false);
                            }}
                            className="rounded-xl border-[var(--border)] bg-[var(--card-bg)] text-[var(--text-soft)] hover:bg-[var(--button-hover)]"
                          >
                            Cancel
                          </Button>
                        </div>
                      </div>
                    ) : null}
                    <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
                      <div className="rounded-2xl border border-[var(--border)] bg-[var(--panel-bg)]/50 p-4">
                        <div className="text-xs uppercase tracking-[0.2em] text-[var(--kpi-title)]">Selected Product</div>
                        <div className="mt-2 text-xl font-semibold text-[var(--text-strong)]">{performanceScorecard.selectedProduct}</div>
                      </div>
                      <div className="rounded-2xl border border-[var(--border)] bg-[var(--panel-bg)]/50 p-4">
                        <div className="text-xs uppercase tracking-[0.2em] text-[var(--kpi-title)]">Working Amount</div>
                        <div className="mt-2 text-xl font-semibold text-[var(--text-strong)]">{currency(performanceScorecard.workingAmount)}</div>
                      </div>
                      <div className="rounded-2xl border border-[var(--border)] bg-[var(--panel-bg)]/50 p-4">
                        <div className="text-xs uppercase tracking-[0.2em] text-[var(--kpi-title)]">{volumeMetric === "gross" ? "Gross Volume" : "Net Volume"}</div>
                        <div className={`mt-2 text-xl font-semibold ${volumeMetric === "gross" ? "text-[var(--text-strong)]" : "text-[var(--kpi-good)]"}`}>{currency(volumeMetric === "gross" ? performanceScorecard.grossVolume : performanceScorecard.netVolume)}</div>
                      </div>
                      <div className="rounded-2xl border border-[var(--border)] bg-[var(--panel-bg)]/50 p-4">
                        <div className="text-xs uppercase tracking-[0.2em] text-[var(--kpi-title)]">Demo %</div>
                        <div className={`mt-2 text-xl font-semibold ${metricAccentClass("Demo %", performanceScorecard.demoPct, "text-[var(--text-strong)]", goalTargets, kpiColorBands)}`}>{pct(performanceScorecard.demoPct)}</div>
                        <div className="mt-1 text-xs text-[var(--kpi-title)]">Goal {pct(goalTargets.demoPct)}</div>
                      </div>
                      <div className="rounded-2xl border border-[var(--border)] bg-[var(--panel-bg)]/50 p-4">
                        <div className="text-xs uppercase tracking-[0.2em] text-[var(--kpi-title)]">Close %</div>
                        <div className={`mt-2 text-xl font-semibold ${metricAccentClass("Close %", performanceScorecard.closePct, "text-[var(--text-strong)]", goalTargets, kpiColorBands)}`}>{pct(performanceScorecard.closePct, 0)}</div>
                        <div className="mt-1 text-xs text-[var(--kpi-title)]">Goal {pct(goalTargets.closePct, 0)}</div>
                      </div>
                      <div className="rounded-2xl border border-[var(--border)] bg-[var(--panel-bg)]/50 p-4">
                        <div className="text-xs uppercase tracking-[0.2em] text-[var(--kpi-title)]">Net %</div>
                        <div className={`mt-2 text-xl font-semibold ${metricAccentClass("Net %", performanceScorecard.netPct, "text-[var(--text-strong)]", goalTargets, kpiColorBands)}`}>{pct(performanceScorecard.netPct)}</div>
                        <div className="mt-1 text-xs text-[var(--kpi-title)]">Goal {pct(goalTargets.netPct)}</div>
                      </div>
                      <div className="rounded-2xl border border-[var(--border)] bg-[var(--panel-bg)]/50 p-4">
                        <div className="text-xs uppercase tracking-[0.2em] text-[var(--kpi-title)]">NSLI</div>
                        <div className={`mt-2 text-xl font-semibold ${metricAccentClass("NSLI", performanceScorecard.nsli, "text-[var(--text-strong)]", goalTargets, kpiColorBands)}`}>{currency(performanceScorecard.nsli)}</div>
                        <div className="mt-1 text-xs text-[var(--kpi-title)]">Goal {currency(goalTargets.nsli)}</div>
                      </div>
                      <div className="rounded-2xl border border-[var(--border)] bg-[var(--panel-bg)]/50 p-4">
                        <div className="text-xs uppercase tracking-[0.2em] text-[var(--kpi-title)]">GSLI</div>
                        <div className="mt-2 text-xl font-semibold text-[var(--text-strong)]">{currency(performanceScorecard.gsli)}</div>
                        <div className="mt-1 text-xs text-[var(--kpi-title)]">Sold Price Split ÷ Issue Split</div>
                      </div>
                    </div>
                  </CardContent>
                </Card>
              </div>

              <div className="mt-6 grid grid-cols-1 gap-4 xl:grid-cols-12">
                <Card className="xl:col-span-12 border-[var(--border-strong)] bg-[var(--card-bg)]">
                  <CardHeader>
                    <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
                      <CardTitle className="text-sm uppercase tracking-[0.25em] text-[#9aa7bb]">Week-over-Week Change</CardTitle>
                      <div className="flex items-center gap-2">
                        <button
                          type="button"
                          onClick={() => setWeekOverWeekMetric("nsli")}
                          className={`rounded-xl border px-3 py-1.5 text-xs font-semibold ${weekOverWeekMetric === "nsli" ? "border-[var(--toggle-active-border)] bg-[var(--toggle-active-bg)] text-[var(--toggle-active-text)]" : "border-[var(--border)] bg-[var(--button-bg)] text-[#a9b4c5] hover:bg-[var(--button-hover)]"}`}
                        >
                          NSLI
                        </button>
                        <button
                          type="button"
                          onClick={() => setWeekOverWeekMetric("volume")}
                          className={`rounded-xl border px-3 py-1.5 text-xs font-semibold ${weekOverWeekMetric === "volume" ? "border-[var(--toggle-active-border)] bg-[var(--toggle-active-bg)] text-[var(--toggle-active-text)]" : "border-[var(--border)] bg-[var(--button-bg)] text-[#a9b4c5] hover:bg-[var(--button-hover)]"}`}
                        >
                          Volume
                        </button>
                      </div>
                    </div>
                  </CardHeader>
                  <CardContent className="h-[440px]">
                    <ResponsiveContainer width="100%" height="100%">
                      <LineChart data={weekChangeData} margin={{ top: 18, right: 40, left: 6, bottom: 68 }}>
                        <CartesianGrid stroke="#1e293b" vertical={false} />
                        <XAxis
                          dataKey="label"
                          stroke="#94a3b8"
                          interval={0}
                          height={58}
                          angle={-33}
                          textAnchor="end"
                          tick={{ fill: "#94a3b8", fontSize: 12 }}
                        />
                        <YAxis stroke="#94a3b8" tick={{ fill: "#94a3b8", fontSize: 12 }} tickFormatter={weekOverWeekAxisFormatter} />
                        <Tooltip contentStyle={tooltipContentStyle} labelStyle={tooltipLabelStyle} itemStyle={tooltipItemStyle} formatter={weekOverWeekTooltipFormatter} />
                        <ReferenceLine y={0} stroke="#f59e0b" strokeDasharray="6 6" ifOverflow="extendDomain" />
                        <Line
                          type="monotone"
                          dataKey="delta"
                          stroke="#38bdf8"
                          strokeWidth={3}
                          dot={{ r: 4, fill: "#38bdf8" }}
                          activeDot={{ r: 6 }}
                        />
                        <Line
                          type="monotone"
                          dataKey="delta"
                          stroke="transparent"
                          strokeWidth={0}
                          dot={false}
                          activeDot={false}
                          legendType="none"
                          isAnimationActive={false}
                        >
                          <LabelList
                            dataKey="actualValue"
                            content={({ x, y, width, value, payload }) => {
                              const numericDelta = safeNum(payload?.delta);
                              const actualValue = safeNum(value);
                              return (
                                <text
                                  x={safeNum(x) + safeNum(width) / 2}
                                  y={numericDelta >= 0 ? safeNum(y) - 10 : safeNum(y) + 18}
                                  fill="var(--chart-label)"
                                  fontSize={14}
                                  fontWeight="700"
                                  textAnchor="middle"
                                >
                                  {weekOverWeekMetric === "netSales"
                                    ? `${Math.round(actualValue)}`
                                    : currency(actualValue)}
                                </text>
                              );
                            }}
                          />
                        </Line>
                      </LineChart>
                    </ResponsiveContainer>
                  </CardContent>
                </Card>
              </div>

              <div className="mt-6 grid grid-cols-1 gap-4 xl:grid-cols-12">
                <Card className="xl:col-span-12 border-[var(--border-strong)] bg-[var(--card-bg)]">
                  <CardHeader>
                    <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
                      <CardTitle className="text-sm uppercase tracking-[0.25em] text-[#9aa7bb]">{productGraphTitle}</CardTitle>
                      <div className="flex items-center gap-2">
                        <button
                          type="button"
                          onClick={() => setProductGraphMetric("nsli")}
                          className={`rounded-xl border px-3 py-1.5 text-xs font-semibold ${productGraphMetric === "nsli" ? "border-[var(--toggle-active-border)] bg-[var(--toggle-active-bg)] text-[var(--toggle-active-text)]" : "border-[var(--border)] bg-[var(--button-bg)] text-[#a9b4c5] hover:bg-[var(--button-hover)]"}`}
                        >
                          NSLI
                        </button>
                        <button
                          type="button"
                          onClick={() => setProductGraphMetric("volume")}
                          className={`rounded-xl border px-3 py-1.5 text-xs font-semibold ${productGraphMetric === "volume" ? "border-[var(--toggle-active-border)] bg-[var(--toggle-active-bg)] text-[var(--toggle-active-text)]" : "border-[var(--border)] bg-[var(--button-bg)] text-[#a9b4c5] hover:bg-[var(--button-hover)]"}`}
                        >
                          Volume
                        </button>
                      </div>
                    </div>
                  </CardHeader>
                  <CardContent className="h-[320px]">
                    <ResponsiveContainer width="100%" height="100%">
                      <BarChart data={productMixData} layout="vertical">
                        <CartesianGrid stroke="#1e293b" horizontal={false} />
                        <XAxis type="number" stroke="#94a3b8" tick={{ fill: "#94a3b8", fontSize: 12 }} tickFormatter={(v) => `$${Math.round(v / 1000)}k`} />
                        <YAxis type="category" dataKey="name" stroke="#94a3b8" tick={{ fill: "#94a3b8", fontSize: 12 }} width={80} />
                        <Tooltip contentStyle={tooltipContentStyle} labelStyle={tooltipLabelStyle} itemStyle={tooltipItemStyle} formatter={(v) => currency(v)} />
                        <Bar
                          dataKey="value"
                          radius={[0, 8, 8, 0]}
                          label={({ x, y, width, height, value }) => (
                            <text
                              x={safeNum(x) + safeNum(width) + 8}
                              y={safeNum(y) + safeNum(height) / 2 + 4}
                              fill="var(--chart-label)"
                              fontSize={14}
                              fontWeight="700"
                              textAnchor="start"
                            >
                              {productGraphMetric === "netSales"
                                ? `${Math.round(safeNum(value))}`
                                : currency(value)}
                            </text>
                          )}
                        >
                          {productMixData.map((item, index) => (
                            <Cell key={item.name} fill={productColors[index % productColors.length]} />
                          ))}
                        </Bar>
                      </BarChart>
                    </ResponsiveContainer>
                  </CardContent>
                </Card>
              </div>

              <div className="mt-6 grid grid-cols-1 gap-4 xl:grid-cols-12">
                <Card className="xl:col-span-12 border-[var(--border-strong)] bg-[var(--card-bg)]">
                  <CardHeader>
                    <div className="flex flex-col gap-3">
                      <div className="flex flex-col gap-3 lg:flex-row lg:items-center lg:justify-between">
                        <CardTitle className="text-sm uppercase tracking-[0.25em] text-[#9aa7bb]">Dynamic KPI Selector and Income Projection</CardTitle>
                        <div className="text-xs text-[var(--kpi-title)]">Projection basis: {projectorRangeLabel}{projected?.baseRep ? ` • ${projected.baseRep}` : ""}</div>
                      </div>
                      <div className="flex flex-col gap-4 xl:flex-row xl:items-start xl:justify-between">
                        <div className="flex flex-wrap items-center gap-3">
                          {["7", "30", "60", "90"].map((option) => {
                            const isActive = projectorTimeframe === option && !projectorRangeEditorOpen;
                            return (
                              <button
                                key={option}
                                type="button"
                                onClick={() => {
                                  const nextRange = buildLookbackRange(option, dateRange.end || datasetMaxDate, datasetMinDate, datasetMaxDate);
                                  setProjectorTimeframe(option);
                                  setProjectorManualRange(nextRange);
                                  setProjectorDraftRange(nextRange);
                                  setProjectorRangeEditorOpen(false);
                                }}
                                className={`min-w-[62px] rounded-xl border px-3 py-1.5 text-[13px] font-semibold ${isActive ? "border-[var(--toggle-active-border)] bg-[var(--toggle-active-bg)] text-[var(--toggle-active-text)]" : "border-[var(--border)] bg-[var(--button-bg)] text-[#a9b4c5] hover:bg-[var(--button-hover)]"}`}
                              >
                                {option}D
                              </button>
                            );
                          })}
                          <button
                            type="button"
                            onClick={() => {
                              setProjectorDraftRange({
                                start: projectorTimeframe === "manual" ? (projectorManualRange.start || dateRange.start || datasetMinDate) : projectorDataset?.minDate || dateRange.start || datasetMinDate,
                                end: projectorTimeframe === "manual" ? (projectorManualRange.end || dateRange.end || datasetMaxDate) : projectorDataset?.maxDate || dateRange.end || datasetMaxDate,
                              });
                              setProjectorRangeEditorOpen(true);
                            }}
                            className={`inline-flex min-w-[92px] items-center justify-center gap-1 rounded-xl border px-3 py-1.5 text-[13px] font-semibold ${projectorRangeEditorOpen || projectorTimeframe === "manual" ? "border-[#5d6f89] bg-[var(--button-active-bg)] text-[var(--text-strong)] ring-1 ring-[#8ea6c8]/55" : "border-[var(--border)] bg-[var(--button-bg)] text-[#a9b4c5] hover:bg-[var(--button-hover)]"}`}
                          >
                            <CalendarDays className="h-4 w-4" />
                            Range
                          </button>
                        </div>
                        <div className="text-xs text-[var(--kpi-title)]">Combined KPI scenario builder</div>
                      </div>
                      {projectorRangeEditorOpen ? (
                        <div className="max-w-[560px] rounded-[28px] border border-[var(--border-strong)] bg-[var(--card-bg)]/95 p-5 shadow-2xl shadow-slate-950/40">
                          <div className="text-xs font-semibold uppercase tracking-[0.25em] text-[#9aa7bb]">Custom Date Range</div>
                          <div className="mt-4 space-y-4">
                            <div>
                              <Label className="mb-2 block text-base text-[var(--text-soft)]">From</Label>
                              <DatePickerField
                                value={projectorDraftRange.start}
                                onChange={(value) => setProjectorDraftRange((current) => ({ ...current, start: value }))}
                                maxDate={projectorDraftRange.end || datasetMaxDate}
                                minDate={datasetMinDate}
                                placeholder="Choose start date"
                              />
                            </div>
                            <div>
                              <Label className="mb-2 block text-base text-[var(--text-soft)]">To</Label>
                              <DatePickerField
                                value={projectorDraftRange.end}
                                onChange={(value) => setProjectorDraftRange((current) => ({ ...current, end: value }))}
                                minDate={projectorDraftRange.start || datasetMinDate}
                                maxDate={datasetMaxDate}
                                placeholder="Choose end date"
                              />
                            </div>
                            <div className="flex items-center gap-3">
                              <Button
                                type="button"
                                onClick={() => {
                                  const appliedRange = {
                                    start: clampIsoDate(projectorDraftRange.start || datasetMinDate, datasetMinDate, datasetMaxDate),
                                    end: clampIsoDate(projectorDraftRange.end || datasetMaxDate, datasetMinDate, datasetMaxDate),
                                  };
                                  setProjectorTimeframe("manual");
                                  setProjectorManualRange(appliedRange);
                                  setProjectorDraftRange(appliedRange);
                                  setProjectorRangeEditorOpen(false);
                                }}
                                className="min-w-[160px] rounded-2xl bg-lime-600 text-lg font-semibold text-slate-950 hover:bg-lime-500"
                              >
                                Apply
                              </Button>
                              <Button
                                type="button"
                                variant="outline"
                                onClick={() => {
                                  setProjectorDraftRange({
                                    start: projectorTimeframe === "manual" ? (projectorManualRange.start || dateRange.start || datasetMinDate) : projectorDataset?.minDate || dateRange.start || datasetMinDate,
                                    end: projectorTimeframe === "manual" ? (projectorManualRange.end || dateRange.end || datasetMaxDate) : projectorDataset?.maxDate || dateRange.end || datasetMaxDate,
                                  });
                                  setProjectorRangeEditorOpen(false);
                                }}
                                className="rounded-2xl border-[var(--border)] bg-[var(--panel-bg)]/60 text-lg text-[var(--text-soft)] hover:bg-slate-800"
                              >
                                Cancel
                              </Button>
                            </div>
                          </div>
                        </div>
                      ) : null}
                    </div>
                  </CardHeader>
                  <CardContent>
                    <div className="grid grid-cols-1 gap-5 xl:grid-cols-12 xl:items-stretch">
                      <div className="xl:col-span-6">
                        <div className={`rounded-2xl border p-4 ${projectionTone.panel}`}>
                          <div className={`flex items-center justify-between text-sm ${projectionTone.label}`}>
                            <span>Combined KPI lift</span>
                            <span className={`font-medium ${projectionTone.value}`}>{aggregateProjectionDelta.toFixed(1)}%</span>
                          </div>
                          <div className="mt-4 grid grid-cols-1 auto-rows-fr gap-4">
                            {[
                              { key: "demoPct", label: "Demo %", formatter: pct },
                              { key: "closePct", label: "Close %", formatter: pct },
                              { key: "netPct", label: "Net %", formatter: pct },
                              { key: "avgTicket", label: "Avg Ticket", formatter: currency },
                            ].map((metric) => {
                              const baseline = projected?.baselineKpis?.[metric.key] ?? 0;
                              const scenario = projected?.projectedKpis?.[metric.key] ?? baseline;
                              const delta = safeNum(projectorAdjustments[metric.key]);
                              return (
                                <div key={metric.key} className="flex min-h-[118px] flex-col justify-between rounded-xl border border-[var(--border)] bg-[var(--panel-bg)]/50 p-4">
                                  <div className="flex items-center justify-between text-sm text-[var(--text-soft)]">
                                    <span>{metric.label}</span>
                                    <span className={delta >= 0 ? "text-lime-300" : "text-orange-300"}>{delta >= 0 ? "+" : ""}{delta}%</span>
                                  </div>
                                  <div className="mt-1 text-xs text-[var(--kpi-title)]">{metric.formatter(baseline)} → {metric.formatter(scenario)}</div>
                                  <div className="mt-3">
                                    <Slider
                                      min={-20}
                                      max={40}
                                      step={1}
                                      value={[safeNum(projectorAdjustments[metric.key])]}
                                      onValueChange={(value) => setProjectorAdjustments((current) => ({ ...current, [metric.key]: safeNum(value[0]) }))}
                                    />
                                  </div>
                                </div>
                              );
                            })}
                          </div>
                        </div>
                      </div>

                      <div className="xl:col-span-6">
                        {projected ? (
                          <div className="mt-[51.75px] grid grid-cols-1 auto-rows-[118px] gap-4">
                            <div className={`flex h-[118px] flex-col justify-between rounded-2xl border bg-[var(--panel-bg)]/55 p-4 ${projectionTone.cardBorder}`}>
                              <div className="text-xs uppercase tracking-[0.2em] text-[var(--kpi-title)]">Baseline Period Volume</div>
                              <div className={`mt-2 text-2xl font-semibold ${aggregateProjectionDelta < 0 ? "text-[var(--kpi-bad)]" : aggregateProjectionDelta > 0 ? "text-[var(--kpi-good)]" : "text-[var(--text-strong)]"}`}>{currency(projected.baselineVolume)}</div>
                              <div className="mt-1 text-xs text-[var(--kpi-title)]">Current period baseline</div>
                            </div>
                            <div className={`flex h-[118px] flex-col justify-between rounded-2xl border bg-[var(--panel-bg)]/55 p-4 ${projectionTone.cardBorder}`}>
                              <div className="text-xs uppercase tracking-[0.2em] text-[var(--kpi-title)]">Projected Monthly Pace</div>
                              <div className={`mt-2 text-2xl font-semibold ${aggregateProjectionDelta < 0 ? "text-[var(--kpi-bad)]" : aggregateProjectionDelta > 0 ? "text-[var(--kpi-good)]" : "text-amber-400"}`}>{currency(projected.monthlyPace)}</div>
                              <div className="mt-1 text-xs text-[var(--kpi-title)]">Tier {projected.monthlyBonusLabel} • Full monthly volume paid at {(bonusForVolume(projected.monthlyPace).rate * 100).toFixed(1)}%</div>
                            </div>
                            <div className={`flex h-[118px] flex-col justify-between rounded-2xl border bg-[var(--panel-bg)]/55 p-4 ${projectionTone.cardBorder}`}>
                              <div className="text-xs uppercase tracking-[0.2em] text-[var(--kpi-title)]">Projected Monthly Bonus</div>
                              <div className={`mt-2 text-2xl font-semibold ${aggregateProjectionDelta < 0 ? "text-[var(--kpi-bad)]" : aggregateProjectionDelta > 0 ? "text-[var(--kpi-good)]" : "text-violet-300"}`}>{currency(projected.monthlyBonusAmount)}</div>
                              <div className="mt-1 text-xs text-[var(--kpi-title)]">Bonus tier {projected.monthlyBonusLabel}</div>
                            </div>
                            <div className={`flex h-[118px] flex-col justify-between rounded-2xl border bg-[var(--panel-bg)]/55 p-4 ${projectionTone.cardBorder}`}>
                              <div className="text-xs uppercase tracking-[0.2em] text-[var(--kpi-title)]">Projected Monthly Income</div>
                              <div className={`mt-2 text-2xl font-semibold ${aggregateProjectionDelta < 0 ? "text-[var(--kpi-bad)]" : aggregateProjectionDelta > 0 ? "text-[var(--kpi-good)]" : "text-violet-400"}`}>{currency(projected.monthlyIncome)}</div>
                              <div className="mt-1 text-xs text-[var(--kpi-title)]">Monthly commission {currency(projected.monthlyCommission)} + monthly bonus {currency(projected.monthlyBonusAmount)}</div>
                            </div>
                          </div>
                        ) : null}
                      </div>
                    </div>
                  </CardContent>
                </Card>
              </div>

              <div className="mt-6 grid grid-cols-1 gap-4 xl:grid-cols-12">
                <Card className="xl:col-span-12 border-[var(--border-strong)] bg-[var(--card-bg)]">
                  <CardHeader><CardTitle className="text-sm uppercase tracking-[0.25em] text-[#9aa7bb]">Comp Assumptions</CardTitle></CardHeader>
                  <CardContent className="space-y-4">
                    <div className="max-w-[520px]">
                      <Label className="mb-2 block text-[var(--text-soft)]">Commission rate</Label>
                      <div className="flex items-center gap-2">
                        <Button
                          type="button"
                          variant="outline"
                          onClick={() => setCommissionRate((current) => Math.max(0, Math.round((current - 0.005) * 1000) / 1000))}
                          className="h-10 w-10 shrink-0 border-[var(--border)] bg-[var(--panel-bg)] px-0 text-[var(--text-strong)] hover:bg-slate-800"
                        >
                          −
                        </Button>
                        <Input
                          type="text"
                          inputMode="decimal"
                          value={`${commissionRate * 100}%`}
                          onChange={(event) => {
                            const raw = String(event.target.value || "").replace(/%/g, "").trim();
                            setCommissionRate(safeNum(raw) / 100);
                          }}
                          className="border-[var(--border)] bg-[var(--panel-bg)] text-[var(--text-strong)] placeholder:text-[var(--kpi-title)]"
                        />
                        <Button
                          type="button"
                          variant="outline"
                          onClick={() => setCommissionRate((current) => Math.round((current + 0.005) * 1000) / 1000)}
                          className="h-10 w-10 shrink-0 border-[var(--border)] bg-[var(--panel-bg)] px-0 text-[var(--text-strong)] hover:bg-slate-800"
                        >
                          +
                        </Button>
                      </div>
                      <div className="mt-2 text-xs text-[var(--kpi-title)]">Adjusts in 0.5% steps</div>
                    </div>
                    <Separator className="bg-slate-800" />
                    <div className="grid grid-cols-1 gap-3 md:grid-cols-3">
                      {bonusTiers.map((tier) => (
                        <div key={tier.label} className="flex items-center justify-between rounded-xl border border-slate-800 bg-[var(--panel-bg)]/50 px-3 py-2 text-sm text-[#9aa7bb]">
                          <span>{tier.min >= 175000 ? `${currency(tier.min)} and above` : `${currency(tier.min)} to ${currency(tier.max)}`}</span>
                          <span>{(tier.rate * 100).toFixed(1)}%</span>
                        </div>
                      ))}
                    </div>
                    <div className="rounded-xl border border-slate-800 bg-[var(--panel-bg)]/50 px-3 py-2 text-xs text-[var(--kpi-title)]">
                      Bonus is a graduated monthly system: once projected monthly net volume reaches a tier, that tier percentage is paid on the full monthly net volume.
                    </div>
                  </CardContent>
                </Card>
              </div>
              <div className="mt-6 grid grid-cols-1 gap-4 xl:grid-cols-12">
                <Card className="xl:col-span-12 border-[var(--border-strong)] bg-[var(--card-bg)]">
                  <CardHeader>
                    <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
                      <div>
                        <CardTitle className="text-sm uppercase tracking-[0.25em] text-[#9aa7bb]">AI Analytics Recommendations</CardTitle>
                        <div className="mt-1 text-xs text-[var(--kpi-title)]">Refresh to regenerate coaching from the current group, rep, product, and date-range selections.</div>
                      </div>
                      <Button
                        type="button"
                        onClick={refreshAnalystRecommendations}
                        className="rounded-xl bg-[var(--button-active-bg)] text-[var(--text-strong)] hover:bg-[#425268]"
                      >
                        Refresh Insights
                      </Button>
                    </div>
                  </CardHeader>
                  <CardContent className="space-y-4">
                    <div className="flex flex-wrap gap-2">
                      <Badge variant="outline" className="border-[var(--border)] bg-[var(--panel-bg)] text-[var(--text-soft)]">Scope: {studioAnalyticsContext.focusLabel}</Badge>
                      <Badge variant="outline" className="border-[var(--border)] bg-[var(--panel-bg)] text-[var(--text-soft)]">Product: {selectedProduct}</Badge>
                      <Badge variant="outline" className="border-[var(--border)] bg-[var(--panel-bg)] text-[var(--text-soft)]">Range: {effectiveRangeLabel}</Badge>
                    </div>
                    <div className="rounded-2xl border border-[var(--border)] bg-[var(--panel-bg)]/60">
                      <ScrollArea className="h-[320px]">
                        <div className="space-y-3 p-4">
                          {analystMessages.length ? analystMessages.map((message, index) => (
                            <div key={`${message.role}-${index}`} className={`rounded-2xl border px-4 py-3 ${message.role === "assistant" ? "border-[var(--border-strong)] bg-[#273140] text-[var(--text-strong)]" : "border-[var(--border)] bg-[#1d2430] text-[var(--text-soft)]"}`}>
                              <div className="mb-1 text-[10px] uppercase tracking-[0.18em] text-[#8ea0b8]">{message.role === "assistant" ? "AI Analyst" : "You"}</div>
                              <div className="whitespace-pre-wrap text-sm leading-6">{message.content}</div>
                            </div>
                          )) : (
                            <div className="rounded-2xl border border-dashed border-[var(--border)] bg-[#1d2430] px-4 py-6 text-sm text-[#8ea0b8]">
                              No recommendations yet. Click <span className="font-semibold text-[var(--text-soft)]">Refresh Insights</span> to generate a coaching read for the current Review Studio selections.
                            </div>
                          )}
                        </div>
                      </ScrollArea>
                    </div>
                  </CardContent>
                </Card>
              </div>          </div>

          <div className={activeDepartment === "Sales Department" && activeWorkspace === "goalManagement" ? "block" : "hidden"}>
            <div className="sticky top-0 z-30 -mx-4 px-4 pb-1.5 backdrop-blur lg:-mx-6 lg:px-6" style={{ backgroundColor: "var(--header-bg)" }}>
              <div className="flex flex-col gap-1.5 xl:flex-row xl:items-start xl:justify-between">
                <div>
                  <h1 className="text-lg font-semibold tracking-[-0.01em] text-[var(--text-strong)]" style={{ fontFamily: "'Rajdhani', 'Inter', sans-serif" }}>Goal Management</h1>
                  <p className="text-[11px] text-[#9ba8bb]">Adjust KPI targets and compare the current filtered performance to those thresholds.</p>
                </div>
                <div className="flex w-full max-w-[260px] flex-col items-stretch gap-1 xl:items-end">
                  <div className="flex items-center justify-end gap-3">
                    <div className="flex items-center gap-2 rounded-xl border border-[var(--border)] bg-[var(--panel-bg)] px-3 py-2">
                      <Sun className="h-4 w-4 text-[var(--text-soft)]" />
                      <Switch checked={isDarkMode} onCheckedChange={setIsDarkMode} />
                      <Moon className="h-4 w-4 text-[var(--text-soft)]" />
                    </div>
                    <label className="inline-flex min-h-[40px] min-w-[160px] cursor-pointer items-center justify-center gap-1.5 rounded-xl border border-[var(--upload-border)] bg-[var(--upload-bg)] px-3 text-[13px] font-semibold text-[var(--upload-text)] shadow-sm transition hover:bg-[var(--upload-hover-bg)] hover:text-[var(--upload-hover-text)]">
                      <Target className="h-3.5 w-3.5" />
                      <span>Upload Report</span>
                      <input type="file" accept=".xlsx,.xls" className="hidden" onChange={onUpload} />
                    </label>
                  </div>
                </div>
              </div>

              <div className="mt-2 flex flex-col gap-1.5 xl:flex-row xl:items-center xl:justify-between">
                <div className="flex flex-wrap items-center gap-3">
                  {["7", "30", "60", "90"].map((option) => {
                    const isActive = performanceTimeframe === option && !dashboardRangeEditorOpen;
                    return (
                      <button
                        key={`goal-${option}`}
                        type="button"
                        onClick={() => {
                          const nextRange = buildLookbackRange(option, dateRange.end || datasetMaxDate, datasetMinDate, datasetMaxDate);
                          setPerformanceTimeframe(option);
                          setDateRange(nextRange);
                          setDashboardDraftRange(nextRange);
                          setDashboardRangeEditorOpen(false);
                        }}
                        className={`min-w-[62px] rounded-xl border px-3 py-1.5 text-[13px] font-semibold ${isActive ? "border-[var(--toggle-active-border)] bg-[var(--toggle-active-bg)] text-[var(--toggle-active-text)]" : "border-[var(--border)] bg-[var(--button-bg)] text-[#a9b4c5] hover:bg-[var(--button-hover)]"}`}
                      >
                        {option}D
                      </button>
                    );
                  })}
                  <button
                    type="button"
                    onClick={() => {
                      setDashboardDraftRange({
                        start: dateRange.start || datasetMinDate,
                        end: dateRange.end || datasetMaxDate,
                      });
                      setDashboardRangeEditorOpen(true);
                    }}
                    className={`inline-flex items-center justify-center gap-2 rounded-xl border px-4 py-2 text-base font-semibold ${dashboardRangeEditorOpen || performanceTimeframe === "manual" ? "border-[#7a8ea9] bg-[var(--button-active-bg)] text-[var(--text-strong)] ring-1 ring-[var(--toggle-active-ring)]/70" : "border-[var(--border)] bg-[var(--panel-bg)] text-[var(--text-soft)] hover:bg-[var(--button-hover)]"}`}
                  >
                    <CalendarDays className="h-4 w-4" />
                    {effectiveRangeLabel}
                  </button>
                  <span className="text-xs uppercase tracking-[0.2em] text-[var(--kpi-title)]">vs current actuals</span>
                </div>
              </div>
            </div>

            {dashboardRangeEditorOpen ? (
              <div className="mt-4 max-w-[560px] rounded-[28px] border border-[var(--border-strong)] bg-[var(--card-bg)]/95 p-5 shadow-2xl shadow-slate-950/40">
                <div className="text-xs font-semibold uppercase tracking-[0.25em] text-[#9aa7bb]">Custom Date Range</div>
                <div className="mt-4 space-y-4">
                  <div>
                    <Label className="mb-2 block text-base text-[var(--text-soft)]">From</Label>
                    <DatePickerField
                      value={dashboardDraftRange.start}
                      onChange={(value) => setDashboardDraftRange((current) => ({ ...current, start: value }))}
                      maxDate={dashboardDraftRange.end || datasetMaxDate}
                      minDate={datasetMinDate}
                      placeholder="Choose start date"
                    />
                  </div>
                  <div>
                    <Label className="mb-2 block text-base text-[var(--text-soft)]">To</Label>
                    <DatePickerField
                      value={dashboardDraftRange.end}
                      onChange={(value) => setDashboardDraftRange((current) => ({ ...current, end: value }))}
                      minDate={dashboardDraftRange.start || datasetMinDate}
                      maxDate={datasetMaxDate}
                      placeholder="Choose end date"
                    />
                  </div>
                  <div className="flex items-center gap-3">
                    <Button
                      type="button"
                      onClick={() => {
                        const appliedRange = {
                          start: clampIsoDate(dashboardDraftRange.start || datasetMinDate, datasetMinDate, datasetMaxDate),
                          end: clampIsoDate(dashboardDraftRange.end || datasetMaxDate, datasetMinDate, datasetMaxDate),
                        };
                        setPerformanceTimeframe("manual");
                        setDateRange(appliedRange);
                        setDashboardDraftRange(appliedRange);
                        setDashboardRangeEditorOpen(false);
                      }}
                      className="min-w-[160px] rounded-2xl bg-lime-600 text-lg font-semibold text-slate-950 hover:bg-lime-500"
                    >
                      Apply
                    </Button>
                    <Button
                      type="button"
                      variant="outline"
                      onClick={() => setDashboardRangeEditorOpen(false)}
                      className="rounded-2xl border-[var(--border)] bg-[var(--panel-bg)]/60 text-lg text-[var(--text-soft)] hover:bg-slate-800"
                    >
                      Cancel
                    </Button>
                  </div>
                </div>
              </div>
            ) : null}

            <div className="mt-6 space-y-4">
              <Card className="border-[var(--border-strong)] bg-[var(--card-bg)]">
                <CardContent className="p-4">
                  <div className="flex flex-col gap-4 lg:flex-row lg:items-start lg:justify-between">
                    <div>
                      <div className="text-xs uppercase tracking-[0.25em] text-[var(--kpi-title)]">Net Volume - Org Total</div>
                      <div className={`mt-2 text-[2.2rem] font-bold ${annualizedNetVolume >= annualVolumeBands.greenMin ? "text-[var(--kpi-good)]" : annualizedNetVolume >= annualVolumeBands.yellowMin ? "text-[var(--kpi-warn)]" : "text-[var(--kpi-bad)]"}`}>{currency(dashboardMetrics.netVolume)}</div>
                      <div className="mt-1 text-sm text-[var(--kpi-title)]">Selected period actual • Annual pace: <span className={`font-semibold ${annualizedNetVolume >= annualVolumeBands.greenMin ? "text-[var(--kpi-good)]" : annualizedNetVolume >= annualVolumeBands.yellowMin ? "text-[var(--kpi-warn)]" : "text-[var(--kpi-bad)]"}`}>{currency(annualizedNetVolume)}</span></div>
                    </div>
                    <div className="flex items-start gap-8">
                      <div>
                        <div className="text-xs uppercase tracking-[0.2em] text-[var(--kpi-title)]">Annualized Target</div>
                        <div className="mt-2 text-2xl font-semibold text-[var(--text-strong)]">{currency(goalTargets.annualNetVolume)}</div>
                        <div className={`mt-1 text-sm font-semibold ${annualizedNetVolume >= annualVolumeBands.greenMin ? "text-[var(--kpi-good)]" : annualizedNetVolume >= annualVolumeBands.yellowMin ? "text-[var(--kpi-warn)]" : "text-[var(--kpi-bad)]"}`}>{annualizedNetVolume - goalTargets.annualNetVolume >= 0 ? "+" : ""}{currency(annualizedNetVolume - goalTargets.annualNetVolume)} gap</div>
                      </div>
                      <div>
                        <div className="text-xs uppercase tracking-[0.2em] text-[var(--kpi-title)]">Annual Goal</div>
                        <div className="mt-2 flex items-center gap-2">
                          <Button type="button" variant="outline" className="h-10 w-10 border-[var(--border)] bg-[var(--panel-bg)] px-0 text-[var(--text-strong)]" onClick={() => setGoalTargets((current) => ({ ...current, annualNetVolume: Math.max(1000000, current.annualNetVolume - 1000000) }))}>−</Button>
                          <Input
                            type="text"
                            value={formatGoalEditorValue(goalTargets.annualNetVolume, "currency")}
                            onChange={(event) => setGoalTargets((current) => ({ ...current, annualNetVolume: Math.max(1000000, parseGoalEditorValue(event.target.value, "currency")) }))}
                            className="h-10 min-w-[120px] rounded-xl border-[var(--border)] bg-[var(--panel-bg)] px-4 py-2 text-center font-semibold text-[var(--text-strong)]"
                          />
                          <Button type="button" variant="outline" className="h-10 w-10 border-[var(--border)] bg-[var(--panel-bg)] px-0 text-[var(--text-strong)]" onClick={() => setGoalTargets((current) => ({ ...current, annualNetVolume: current.annualNetVolume + 1000000 }))}>+</Button>
                        </div>
                      </div>
                    </div>
                  </div>
                  <div className="mt-4 grid grid-cols-1 gap-3 sm:grid-cols-2 lg:max-w-[520px]">
                    <div>
                      <div className="text-[10px] uppercase tracking-[0.2em] text-[var(--kpi-title)]">Green floor</div>
                      <div className="mt-2 flex items-center gap-2">
                        <Button
                          type="button"
                          variant="outline"
                          className="h-8 w-8 border-[var(--border)] bg-[var(--panel-bg)] px-0 text-[var(--text-strong)]"
                          onClick={() => setAnnualVolumeBands((current) => ({
                            ...current,
                            greenMin: Math.max(current.yellowMin + 500000, current.greenMin - 500000),
                          }))}
                        >
                          −
                        </Button>
                        <Input
                          type="text"
                          value={formatGoalEditorValue(annualVolumeBands.greenMin, "currency")}
                          onChange={(event) => setAnnualVolumeBands((current) => ({
                            ...current,
                            greenMin: Math.max(current.yellowMin + 500000, parseGoalEditorValue(event.target.value, "currency")),
                          }))}
                          className="h-8 min-w-[120px] rounded-xl border-[var(--border)] bg-[var(--panel-bg)] px-3 py-2 text-center font-semibold text-[var(--kpi-good)]"
                        />
                        <Button
                          type="button"
                          variant="outline"
                          className="h-8 w-8 border-[var(--border)] bg-[var(--panel-bg)] px-0 text-[var(--text-strong)]"
                          onClick={() => setAnnualVolumeBands((current) => ({
                            ...current,
                            greenMin: current.greenMin + 500000,
                          }))}
                        >
                          +
                        </Button>
                      </div>
                    </div>
                    <div>
                      <div className="text-[10px] uppercase tracking-[0.2em] text-[var(--kpi-title)]">Yellow floor</div>
                      <div className="mt-2 flex items-center gap-2">
                        <Button
                          type="button"
                          variant="outline"
                          className="h-8 w-8 border-[var(--border)] bg-[var(--panel-bg)] px-0 text-[var(--text-strong)]"
                          onClick={() => setAnnualVolumeBands((current) => ({
                            ...current,
                            yellowMin: Math.max(0, current.yellowMin - 500000),
                          }))}
                        >
                          −
                        </Button>
                        <Input
                          type="text"
                          value={formatGoalEditorValue(annualVolumeBands.yellowMin, "currency")}
                          onChange={(event) => setAnnualVolumeBands((current) => ({
                            ...current,
                            yellowMin: Math.min(current.greenMin - 500000, Math.max(0, parseGoalEditorValue(event.target.value, "currency"))),
                          }))}
                          className="h-8 min-w-[120px] rounded-xl border-[var(--border)] bg-[var(--panel-bg)] px-3 py-2 text-center font-semibold text-[var(--kpi-warn)]"
                        />
                        <Button
                          type="button"
                          variant="outline"
                          className="h-8 w-8 border-[var(--border)] bg-[var(--panel-bg)] px-0 text-[var(--text-strong)]"
                          onClick={() => setAnnualVolumeBands((current) => ({
                            ...current,
                            yellowMin: Math.min(current.greenMin - 500000, current.yellowMin + 500000),
                          }))}
                        >
                          +
                        </Button>
                      </div>
                    </div>
                  </div>
                  <div className="mt-4 h-2.5 rounded-full bg-[var(--panel-bg)]">
                    <div
                      className={`h-2.5 rounded-full ${annualizedNetVolume >= annualVolumeBands.greenMin ? "bg-[var(--kpi-good)]" : annualizedNetVolume >= annualVolumeBands.yellowMin ? "bg-[var(--kpi-warn)]" : "bg-[var(--kpi-bad)]"}`}
                      style={{ width: `${Math.min(100, Math.max(6, (annualizedNetVolume / Math.max(1, goalTargets.annualNetVolume)) * 100))}%` }}
                    />
                  </div>
                </CardContent>
              </Card>

              <div className="grid grid-cols-1 gap-4 md:grid-cols-2 xl:grid-cols-3">
                {goalManagementCards.map((item) => {
                  const gap = safeNum(item.actual) - safeNum(item.goal);
                  const tierConfig = kpiColorBands[item.key] || null;
                  const isGood = tierConfig ? safeNum(item.actual) >= safeNum(tierConfig.greenMin) : safeNum(item.actual) >= safeNum(item.goal);
                  const isWarn = tierConfig
                    ? safeNum(item.actual) >= safeNum(tierConfig.yellowMin) && safeNum(item.actual) < safeNum(tierConfig.greenMin)
                    : !isGood && safeNum(item.actual) >= safeNum(item.goal) * 0.95;
                  const toneClass = isGood ? "text-[var(--kpi-good)]" : isWarn ? "text-[var(--kpi-warn)]" : "text-[var(--kpi-bad)]";
                  const barClass = isGood ? "bg-[var(--kpi-good)]" : isWarn ? "bg-[var(--kpi-warn)]" : "bg-[var(--kpi-bad)]";
                  return (
                    <Card key={item.key} className="border-[var(--border-strong)] bg-[var(--card-bg)]">
                      <CardContent className="p-4">
                        <div className="flex items-start justify-between gap-4">
                          <div>
                            <div className="text-xs uppercase tracking-[0.2em] text-[var(--kpi-title)]">{item.title}</div>
                            <div className="mt-1 text-xs text-[var(--kpi-title)]">{item.subtitle}</div>
                          </div>
                          <div className="text-right">
                            <div className="text-xs text-[var(--kpi-title)]">Actual</div>
                            <div className={`mt-1 text-[1.9rem] font-bold ${toneClass}`}>{item.formatter(item.actual)}</div>
                          </div>
                        </div>
                        <div className="mt-5 flex items-center justify-between gap-3">
                          <div>
                            <div className="text-[10px] uppercase tracking-[0.2em] text-[var(--kpi-title)]">Goal</div>
                            <div className="mt-2 flex items-center gap-2">
                              <Button type="button" variant="outline" className="h-8 w-8 border-[var(--border)] bg-[var(--panel-bg)] px-0 text-[var(--text-strong)]" onClick={() => setGoalTargets((current) => ({ ...current, [item.key]: Math.max(item.step, safeNum(current[item.key]) - item.step) }))}>−</Button>
                              <Input
                                  type="text"
                                  value={formatGoalEditorValue(goalTargets[item.key], item.inputType || "currency")}
                                  onChange={(event) => setGoalTargets((current) => ({ ...current, [item.key]: Math.max(item.step, parseGoalEditorValue(event.target.value, item.inputType || "currency")) }))}
                                  className="h-8 min-w-[88px] rounded-xl border-[var(--border)] bg-[var(--panel-bg)] px-3 py-2 text-center font-semibold text-[var(--text-strong)]"
                                />
                              <Button type="button" variant="outline" className="h-8 w-8 border-[var(--border)] bg-[var(--panel-bg)] px-0 text-[var(--text-strong)]" onClick={() => setGoalTargets((current) => ({ ...current, [item.key]: safeNum(current[item.key]) + item.step }))}>+</Button>
                            </div>
                          </div>
                          <div className="text-right">
                            <div className="text-[10px] uppercase tracking-[0.2em] text-[var(--kpi-title)]">Gap</div>
                            <div className={`mt-2 text-xl font-semibold ${toneClass}`}>{gap >= 0 ? "+" : ""}{item.key.includes("Pct") ? pct(gap, 0) : item.formatter(gap)}</div>
                          </div>
                        </div>
                        {tierConfig ? (
                          <div className="mt-4 grid grid-cols-1 gap-3 sm:grid-cols-2">
                            <div>
                              <div className="text-[10px] uppercase tracking-[0.2em] text-[var(--kpi-title)]">Green floor</div>
                              <div className="mt-2 flex items-center gap-2">
                                <Button
                                  type="button"
                                  variant="outline"
                                  className="h-8 w-8 border-[var(--border)] bg-[var(--panel-bg)] px-0 text-[var(--text-strong)]"
                                  onClick={() => setKpiColorBands((current) => ({
                                    ...current,
                                    [item.key]: {
                                      ...current[item.key],
                                      greenMin: Math.max(safeNum(current[item.key]?.yellowMin) + item.step, safeNum(current[item.key]?.greenMin) - item.step),
                                    },
                                  }))}
                                >
                                  −
                                </Button>
                                <Input
                                  type="text"
                                  value={formatGoalEditorValue(tierConfig.greenMin, item.inputType || "currency")}
                                  onChange={(event) => setKpiColorBands((current) => ({
                                    ...current,
                                    [item.key]: {
                                      ...current[item.key],
                                      greenMin: Math.max(safeNum(current[item.key]?.yellowMin) + item.step, parseGoalEditorValue(event.target.value, item.inputType || "currency")),
                                    },
                                  }))}
                                  className="h-8 min-w-[88px] rounded-xl border-[var(--border)] bg-[var(--panel-bg)] px-3 py-2 text-center font-semibold text-[var(--kpi-good)]"
                                />
                                <Button
                                  type="button"
                                  variant="outline"
                                  className="h-8 w-8 border-[var(--border)] bg-[var(--panel-bg)] px-0 text-[var(--text-strong)]"
                                  onClick={() => setKpiColorBands((current) => ({
                                    ...current,
                                    [item.key]: {
                                      ...current[item.key],
                                      greenMin: safeNum(current[item.key]?.greenMin) + item.step,
                                    },
                                  }))}
                                >
                                  +
                                </Button>
                              </div>
                            </div>
                            <div>
                              <div className="text-[10px] uppercase tracking-[0.2em] text-[var(--kpi-title)]">Yellow floor</div>
                              <div className="mt-2 flex items-center gap-2">
                                <Button
                                  type="button"
                                  variant="outline"
                                  className="h-8 w-8 border-[var(--border)] bg-[var(--panel-bg)] px-0 text-[var(--text-strong)]"
                                  onClick={() => setKpiColorBands((current) => ({
                                    ...current,
                                    [item.key]: {
                                      ...current[item.key],
                                      yellowMin: Math.max(0, safeNum(current[item.key]?.yellowMin) - item.step),
                                    },
                                  }))}
                                >
                                  −
                                </Button>
                                <Input
                                  type="text"
                                  value={formatGoalEditorValue(tierConfig.yellowMin, item.inputType || "currency")}
                                  onChange={(event) => setKpiColorBands((current) => ({
                                    ...current,
                                    [item.key]: {
                                      ...current[item.key],
                                      yellowMin: Math.min(safeNum(current[item.key]?.greenMin) - item.step, Math.max(0, parseGoalEditorValue(event.target.value, item.inputType || "currency"))),
                                    },
                                  }))}
                                  className="h-8 min-w-[88px] rounded-xl border-[var(--border)] bg-[var(--panel-bg)] px-3 py-2 text-center font-semibold text-[var(--kpi-warn)]"
                                />
                                <Button
                                  type="button"
                                  variant="outline"
                                  className="h-8 w-8 border-[var(--border)] bg-[var(--panel-bg)] px-0 text-[var(--text-strong)]"
                                  onClick={() => setKpiColorBands((current) => ({
                                    ...current,
                                    [item.key]: {
                                      ...current[item.key],
                                      yellowMin: Math.min(safeNum(current[item.key]?.greenMin) - item.step, safeNum(current[item.key]?.yellowMin) + item.step),
                                    },
                                  }))}
                                >
                                  +
                                </Button>
                              </div>
                            </div>
                            <div className="sm:col-span-2 text-[10px] text-[var(--kpi-title)]">Red applies to any value below the yellow floor.</div>
                          </div>
                        ) : null}
                        <div className="mt-4 h-1.5 rounded-full bg-[var(--panel-bg)]">
                          <div className={`${barClass} h-1.5 rounded-full`} style={{ width: `${Math.min(100, Math.max(5, (safeNum(item.actual) / Math.max(safeNum(item.goal), 0.0001)) * 100))}%` }} />
                        </div>
                      </CardContent>
                    </Card>
                  );
                })}
              </div>
            </div>
          </div>

          <div className={activeDepartment === "Sales Department" && activeWorkspace === "tierManagement" ? "block" : "hidden"}>
            <div className="sticky top-0 z-30 -mx-4 px-4 pb-1.5 backdrop-blur lg:-mx-6 lg:px-6" style={{ backgroundColor: "var(--header-bg)" }}>
              <div className="flex flex-col gap-2 xl:flex-row xl:items-start xl:justify-between">
                <div>
                  <h1 className="text-lg font-semibold tracking-[-0.01em] text-[var(--text-strong)]" style={{ fontFamily: "'Rajdhani', 'Inter', sans-serif" }}>Tier Management</h1>
                  <p className="text-[11px] text-[#9ba8bb]">Close % thresholds determine A, B, or C classification. Changes propagate live across the movement view.</p>
                </div>
                <div className="flex items-center gap-2">
                  <Button
                    type="button"
                    variant="outline"
                    onClick={() => {
                      setTierThresholds(DEFAULT_TIER_THRESHOLDS);
                      setMovementConfig(DEFAULT_MOVEMENT_CONFIG);
                    }}
                    className="border-[var(--border)] bg-[var(--panel-bg)] text-[var(--text-soft)] hover:bg-[var(--button-hover)]"
                  >
                    Reset
                  </Button>
                </div>
              </div>
            </div>

            <div className="mt-6 space-y-5">
              <div className="grid grid-cols-1 gap-4 lg:grid-cols-3">
                {[
                  { key: "A", title: "A", subtitle: "Elite performers", threshold: tierThresholds.aMin, count: tierManagementData.distribution.A, tone: "border-emerald-500/50", textTone: "text-emerald-400", editable: true },
                  { key: "B", title: "B", subtitle: "Core performers", threshold: tierThresholds.bMin, count: tierManagementData.distribution.B, tone: "border-amber-500/50", textTone: "text-amber-400", editable: true },
                  { key: "C", title: "C", subtitle: "Needs development", threshold: tierThresholds.bMin, count: tierManagementData.distribution.C, tone: "border-rose-500/50", textTone: "text-rose-400", editable: false },
                ].map((card) => (
                  <Card key={card.key} className={`border bg-[var(--card-bg)] ${card.tone}`}>
                    <CardContent className="p-4">
                      <div className="flex items-start justify-between gap-3">
                        <div>
                          <div className={`inline-flex h-7 min-w-[28px] items-center justify-center rounded-md border px-2 text-sm font-bold ${card.tone} ${card.textTone}`}>{card.title}</div>
                          <div className="mt-2 text-sm text-[var(--kpi-title)]">{card.subtitle}</div>
                        </div>
                        <div className="text-xs text-[var(--kpi-title)]">{card.count} reps</div>
                      </div>
                      <div className="mt-5 flex items-center gap-2">
                        {card.editable ? (
                          <>
                            <Button type="button" variant="outline" className="h-8 w-8 border-[var(--border)] bg-[var(--panel-bg)] px-0 text-[var(--text-strong)]" onClick={() => setTierThresholds((current) => card.key === "A" ? { ...current, aMin: Math.max(current.bMin + 0.01, Math.round((current.aMin - 0.01) * 100) / 100) } : { ...current, bMin: Math.max(0.1, Math.min(current.aMin - 0.01, Math.round((current.bMin - 0.01) * 100) / 100)) })}>−</Button>
                            <div className="min-w-[88px] rounded-xl border border-[var(--border)] bg-[var(--panel-bg)] px-3 py-2 text-center font-semibold text-[var(--text-strong)]">{pct(card.threshold, 0)}</div>
                            <Button type="button" variant="outline" className="h-8 w-8 border-[var(--border)] bg-[var(--panel-bg)] px-0 text-[var(--text-strong)]" onClick={() => setTierThresholds((current) => card.key === "A" ? { ...current, aMin: Math.min(0.6, Math.round((current.aMin + 0.01) * 100) / 100) } : { ...current, bMin: Math.min(current.aMin - 0.01, Math.round((current.bMin + 0.01) * 100) / 100) })}>+</Button>
                          </>
                        ) : (
                          <div className="text-2xl font-semibold text-[var(--text-strong)]">{`< ${pct(tierThresholds.bMin, 0)}`}</div>
                        )}
                      </div>
                    </CardContent>
                  </Card>
                ))}
              </div>

              <div className="flex flex-wrap items-center gap-3 text-sm">
                <span className="text-[var(--kpi-title)]">Live dist.</span>
                {[
                  { key: "A", count: tierManagementData.distribution.A, text: "text-emerald-400" },
                  { key: "B", count: tierManagementData.distribution.B, text: "text-amber-400" },
                  { key: "C", count: tierManagementData.distribution.C, text: "text-rose-400" },
                ].map((item) => (
                  <div key={item.key} className="flex items-center gap-2">
                    <div className={`h-4 w-[92px] rounded-sm ${item.key === "A" ? "bg-emerald-500" : item.key === "B" ? "bg-amber-500" : "bg-rose-500"}`} />
                    <span className={`font-semibold ${item.text}`}>{item.key}: {item.count} ({tierManagementData.distribution.total ? Math.round((item.count / tierManagementData.distribution.total) * 100) : 0}%)</span>
                  </div>
                ))}
              </div>

              <div>
                <div className="text-xl font-semibold text-[var(--text-strong)]">Movement Matrix</div>
                <div className="mt-1 text-sm text-[var(--kpi-title)]">90D tier baseline vs 30D tier current. Adjust lead multiplier and guidance per bucket.</div>
              </div>

              <div className="space-y-3">
                {[0, 1, 2].map((rowIndex) => {
                  const rowBuckets = tierManagementData.buckets.slice(rowIndex * 3, rowIndex * 3 + 3);
                  return (
                    <div key={`row-${rowIndex}`} className="space-y-3">
                      <div className="grid grid-cols-3 gap-3 rounded-xl border border-[var(--border)] bg-[var(--panel-bg)] px-4 py-2 text-center text-sm text-[var(--kpi-title)]">
                        {rowBuckets.map((bucket) => (
                          <div key={`${bucket.bucketId}-header`}>{`90D:${bucket.bucketId[0]} / 30D:${bucket.bucketId[1]}`}</div>
                        ))}
                      </div>
                      <div className="grid grid-cols-1 gap-3 xl:grid-cols-3">
                        {rowBuckets.map((bucket) => (
                          <Card key={bucket.bucketId} className={`border bg-[var(--card-bg)] ${bucket.tone}`}>
                            <CardContent className="p-4">
                              <div className={`text-lg font-semibold ${bucket.titleTone}`}>{bucket.title}</div>
                              <div className="mt-1 text-xs text-[var(--kpi-title)]">{bucket.summary} | {bucket.count} reps</div>
                              <div className="mt-4 grid grid-cols-2 gap-4">
                                <div>
                                  <div className="text-[10px] uppercase tracking-[0.18em] text-[var(--kpi-title)]">Lead Mult</div>
                                  <div className="mt-2 flex items-center gap-2">
                                    <Button type="button" variant="outline" className="h-8 w-8 border-[var(--border)] bg-[var(--panel-bg)] px-0 text-[var(--text-strong)]" onClick={() => setMovementConfig((current) => ({ ...current, [bucket.bucketId]: { ...current[bucket.bucketId], multiplier: Math.max(0.5, Math.round((safeNum(current[bucket.bucketId]?.multiplier) - 0.05) * 100) / 100) } }))}>−</Button>
                                    <div className="min-w-[76px] rounded-lg border border-[var(--border)] bg-[var(--panel-bg)] px-3 py-2 text-center font-semibold text-[var(--text-strong)]">{safeNum(bucket.config.multiplier).toFixed(2)}x</div>
                                    <Button type="button" variant="outline" className="h-8 w-8 border-[var(--border)] bg-[var(--panel-bg)] px-0 text-[var(--text-strong)]" onClick={() => setMovementConfig((current) => ({ ...current, [bucket.bucketId]: { ...current[bucket.bucketId], multiplier: Math.min(1.5, Math.round((safeNum(current[bucket.bucketId]?.multiplier) + 0.05) * 100) / 100) } }))}>+</Button>
                                  </div>
                                </div>
                                <div>
                                  <div className="text-[10px] uppercase tracking-[0.18em] text-[var(--kpi-title)]">Guidance</div>
                                  <Select value={bucket.config.guidance} onValueChange={(value) => setMovementConfig((current) => ({ ...current, [bucket.bucketId]: { ...current[bucket.bucketId], guidance: value } }))}>
                                    <SelectTrigger className="mt-2 h-10 border-[var(--border)] bg-[var(--panel-bg)] text-sm text-[var(--text-strong)]">
                                      <SelectValue />
                                    </SelectTrigger>
                                    <SelectContent>
                                      {[
                                        "Prioritize",
                                        "Maintain",
                                        "Cap / coach",
                                      ].map((option) => (
                                        <SelectItem key={`${bucket.bucketId}-${option}`} value={option}>{option}</SelectItem>
                                      ))}
                                    </SelectContent>
                                  </Select>
                                </div>
                              </div>
                            </CardContent>
                          </Card>
                        ))}
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          </div>

          <div className={activeDepartment === "Sales Department" && activeWorkspace === "teamManagement" ? "block" : "hidden"}>
            <div className="sticky top-0 z-30 -mx-4 px-4 pb-1.5 backdrop-blur lg:-mx-6 lg:px-6" style={{ backgroundColor: "var(--header-bg)" }}>
              <div className="flex flex-col gap-2 xl:flex-row xl:items-start xl:justify-between">
                <div>
                  <h1 className="text-lg font-semibold tracking-[-0.01em] text-[var(--text-strong)]" style={{ fontFamily: "'Rajdhani', 'Inter', sans-serif" }}>Team Management</h1>
                  <p className="text-[11px] text-[#9ba8bb]">Reassign reps across teams. Changes apply live across filters and review screens.</p>
                </div>
                <div className="flex items-center gap-2">
                  <Button type="button" onClick={createManagedTeam} className="bg-lime-600 text-slate-950 hover:bg-lime-500">+ New Team</Button>
                  <Button type="button" variant="outline" onClick={() => setCustomGroups({})} className="border-[var(--border)] bg-[var(--panel-bg)] text-[var(--text-soft)] hover:bg-[var(--button-hover)]">Reset</Button>
                </div>
              </div>
            </div>

            <div className="mt-6 grid grid-cols-1 gap-4 xl:grid-cols-4">
              {teamManagementColumns.map((team) => (
                <Card key={team.name} className="border-[var(--border-strong)] bg-[var(--card-bg)]">
                  <CardHeader className={`rounded-t-xl border-b ${team.theme.header}`}>
                    <div className="flex items-center justify-between gap-3">
                      <div>
                        <CardTitle className="text-lg font-semibold tracking-[0.04em]">{team.name}</CardTitle>
                        <div className="mt-1 text-xs opacity-80">{team.reps.length} reps</div>
                      </div>
                    </div>
                  </CardHeader>
                  <CardContent className="p-3">
                    <div className="space-y-3">
                      {team.reps.length ? team.reps.map((rep) => {
                        const tier = repTierLetter(rep);
                        const health = repHealthValue(rep);
                        return (
                          <div key={`${team.name}-${rep.rep}`} className="rounded-xl border border-[var(--border)] bg-[var(--panel-bg)] p-3">
                            <div className="flex items-start justify-between gap-3">
                              <div>
                                <div className="font-semibold text-[var(--text-strong)]">{rep.rep}</div>
                                <div className="mt-1 text-xs text-[var(--kpi-title)]">{rep.movement}</div>
                              </div>
                              <div className="flex items-center gap-2">
                                <span className={`inline-flex h-7 w-7 items-center justify-center rounded-md border text-xs font-bold ${tier === "A" ? "border-emerald-500/40 bg-emerald-500/10 text-emerald-300" : tier === "B" ? "border-amber-500/40 bg-amber-500/10 text-amber-300" : "border-rose-500/40 bg-rose-500/10 text-rose-300"}`}>{tier}</span>
                                <span className={`text-sm font-semibold ${health >= 0.75 ? "text-emerald-300" : health >= 0.6 ? "text-amber-300" : "text-rose-300"}`}>{health.toFixed(2)}</span>
                              </div>
                            </div>
                            <div className="mt-3 grid grid-cols-[1fr_120px] items-center gap-2">
                              <div className="text-xs text-[var(--kpi-title)]">{currency(rep.netVolume)} net volume • {pct(rep.closePct, 0)} close</div>
                              <Select value={team.name} onValueChange={(value) => moveRepToTeam(rep.rep, value)}>
                                <SelectTrigger className="h-8 border-[var(--border)] bg-[var(--panel-bg)] text-xs text-[var(--text-soft)]">
                                  <SelectValue />
                                </SelectTrigger>
                                <SelectContent>
                                  {manageableTeamNames.map((teamOption) => (
                                    <SelectItem key={`${rep.rep}-${teamOption}`} value={teamOption}>{teamOption}</SelectItem>
                                  ))}
                                </SelectContent>
                              </Select>
                            </div>
                          </div>
                        );
                      }) : (
                        <div className="flex h-[180px] items-center justify-center rounded-xl border border-dashed border-[var(--border)] bg-[var(--panel-bg)] text-sm text-[var(--kpi-title)]">Drop reps here</div>
                      )}
                    </div>
                  </CardContent>
                </Card>
              ))}
            </div>
          </div>
        </main>
      </div>
    </div>
  );
}
