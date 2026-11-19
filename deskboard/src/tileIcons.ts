import { useEffect, useState } from "react";
import { api } from "./api";
import type { Tile } from "./types";

/** 타일에 그릴 그림을 어디서 구할지. 이모지를 쓰거나 웹 주소인 타일은 구할 그림이 없다. */
function iconSource(tile: Tile): { path: string; thumbnail: boolean } | null {
  if (tile.icon.startsWith("emoji:")) return null;
  if (tile.icon.startsWith("file:")) return { path: tile.icon.slice(5), thumbnail: true };
  return tile.kind === "link" ? null : { path: tile.target, thumbnail: false };
}

const keyOf = (source: { path: string; thumbnail: boolean }) => `${source.thumbnail ? "T" : "I"}|${source.path}`;

// 같은 대상의 아이콘은 한 번만 물어본다. (값이 null 이면 물어봤지만 없었던 것)
const loaded = new Map<string, string | null>();
const asked = new Set<string>();

/** 타일들의 아이콘을 구해 두고, 타일을 주면 그 아이콘(없으면 null)을 돌려주는 함수를 준다. */
export function useTileIcons(tiles: Tile[]): (tile: Tile) => string | null {
  const [, refresh] = useState(0);
  const wanted = tiles.map(iconSource).filter((s) => s !== null);
  const signature = wanted.map(keyOf).join("\n");

  useEffect(() => {
    for (const source of wanted) {
      const key = keyOf(source);
      if (asked.has(key)) continue;
      asked.add(key);
      api.tileIcon(source.path, source.thumbnail).then(
        (url) => {
          loaded.set(key, url);
          refresh((n) => n + 1);
        },
        () => loaded.set(key, null),
      );
    }
  }, [signature]); // eslint-disable-line react-hooks/exhaustive-deps

  return (tile) => {
    const source = iconSource(tile);
    return source ? (loaded.get(keyOf(source)) ?? null) : null;
  };
}
