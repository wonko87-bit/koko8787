import type { ReactNode } from "react";
import type { TileKind } from "../types";

function Svg({ children, size = 16 }: { children: ReactNode; size?: number }) {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="2"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
    >
      {children}
    </svg>
  );
}

const KIND_PATHS: Record<TileKind, ReactNode> = {
  folder: <path d="M3 7a2 2 0 0 1 2-2h4l2 2.5h8a2 2 0 0 1 2 2V17a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2Z" />,
  program: (
    <>
      <rect x="3" y="4.5" width="18" height="15" rx="2.5" />
      <path d="M3 9h18M8 13.5l2.5 2-2.5 2" />
    </>
  ),
  file: (
    <>
      <path d="M6 3h8l4 4v12a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2Z" />
      <path d="M14 3v4h4M8 13h8M8 17h5" />
    </>
  ),
  link: (
    <>
      <circle cx="12" cy="12" r="9" />
      <path d="M3 12h18M12 3c2.8 3 2.8 15 0 18M12 3c-2.8 3-2.8 15 0 18" />
    </>
  ),
};

export function KindIcon({ kind, size }: { kind: TileKind; size: number }) {
  return <Svg size={size}>{KIND_PATHS[kind]}</Svg>;
}

export const PlusIcon = () => (
  <Svg size={14}>
    <path d="M12 5v14M5 12h14" />
  </Svg>
);

export const CloseIcon = ({ size = 14 }: { size?: number }) => (
  <Svg size={size}>
    <path d="M6 6l12 12M18 6 6 18" />
  </Svg>
);

export const MoreIcon = () => (
  <Svg size={14}>
    <circle cx="5" cy="12" r="1" />
    <circle cx="12" cy="12" r="1" />
    <circle cx="19" cy="12" r="1" />
  </Svg>
);

export const LockIcon = ({ open }: { open: boolean }) => (
  <Svg>
    <rect x="5" y="11" width="14" height="9" rx="2" />
    <path d={open ? "M8 11V8a4 4 0 0 1 7.5-2" : "M8 11V8a4 4 0 0 1 8 0v3"} />
  </Svg>
);

export const SlidersIcon = () => (
  <Svg>
    <path d="M4 7h9M17 7h3M4 17h3M11 17h9" />
    <circle cx="15" cy="7" r="2" />
    <circle cx="9" cy="17" r="2" />
  </Svg>
);

export const ChecklistIcon = () => (
  <Svg size={13}>
    <path d="M3.5 6.5l2 2 3.5-4M3.5 16.5l2 2 3.5-4M13 7h8M13 17h8" />
  </Svg>
);

export const TextIcon = () => (
  <Svg size={13}>
    <path d="M4 6h16M4 12h16M4 18h10" />
  </Svg>
);
