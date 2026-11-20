import { useRef, useState } from "react";
import type { MouseEvent, PointerEvent as ReactPointerEvent } from "react";
import type { Category, Tile, TileSize } from "../types";
import { KindIcon, MoreIcon, PlusIcon } from "./icons";

const GLYPH_SIZE: Record<TileSize, number> = { s: 16, m: 20, l: 26 };
/** 이만큼 움직이기 전에는 끌기가 아니라 클릭으로 본다. */
const DRAG_THRESHOLD = 6;

interface Props {
  categories: Category[];
  tileSize: TileSize;
  iconFor: (tile: Tile) => string | null;
  /** 탐색기에서 끌어온 파일이 놓일 카테고리 */
  fileDropCategoryId: string | null;
  /** 찾는 말. 비어 있지 않으면 맞는 타일만 보인다. */
  query: string;
  onLaunch: (tile: Tile) => void;
  onTileMenu: (at: { x: number; y: number }, tile: Tile, categoryId: string) => void;
  onCategoryMenu: (at: { x: number; y: number }, category: Category) => void;
  onAddTile: (categoryId: string) => void;
  onAddCategory: () => void;
  onMoveTile: (tileId: string, toCategoryId: string, beforeTileId: string | null) => void;
}

/** 끌고 있는 타일을 놓으면 들어갈 자리 */
interface DropSpot {
  categoryId: string;
  beforeTileId: string | null;
}

/** 화면의 한 점이 어느 카테고리의 어느 자리인지. 타일의 왼쪽 절반이면 그 앞, 오른쪽 절반이면 그 뒤다. */
function spotAt(x: number, y: number): DropSpot | null {
  const under = document.elementFromPoint(x, y);
  const categoryId = under?.closest<HTMLElement>("[data-category-id]")?.dataset.categoryId;
  if (!categoryId) return null;
  const tile = under?.closest<HTMLElement>("[data-tile-id]");
  if (!tile) return { categoryId, beforeTileId: null };
  const box = tile.getBoundingClientRect();
  if (x <= box.left + box.width / 2) return { categoryId, beforeTileId: tile.dataset.tileId ?? null };
  const next = tile.nextElementSibling as HTMLElement | null;
  return { categoryId, beforeTileId: next?.dataset.tileId ?? null };
}

/** 타일 이름이나 대상의 마지막 이름(파일·폴더 이름)에 찾는 말이 들어 있는가. 대소문자는 가리지 않는다. */
export function matchesTile(tile: Tile, query: string): boolean {
  const wanted = query.trim().toLowerCase();
  if (!wanted) return true;
  const leaf = tile.target.replace(/[\\/]+$/, "").split(/[\\/]/).pop() ?? "";
  return tile.name.toLowerCase().includes(wanted) || leaf.toLowerCase().includes(wanted);
}

function TileFace({ tile, icon, glyphSize }: { tile: Tile; icon: string | null; glyphSize: number }) {
  if (tile.icon.startsWith("emoji:")) return <span className="tile-emoji">{tile.icon.slice(6)}</span>;
  if (icon) return <img className="tile-image" src={icon} alt="" draggable={false} />;
  return (
    <span className={`tile-glyph kind-${tile.kind}`}>
      <KindIcon kind={tile.kind} size={glyphSize} />
    </span>
  );
}

export function Shortcuts(props: Props) {
  const { tileSize, iconFor, fileDropCategoryId, query } = props;
  const searching = query.trim() !== "";
  // 찾는 중에는 맞는 타일만, 그런 타일이 있는 카테고리만 남긴다. 맨 처음 것이 Enter 로 열린다.
  const categories = searching
    ? props.categories.map((c) => ({ ...c, tiles: c.tiles.filter((t) => matchesTile(t, query)) })).filter((c) => c.tiles.length > 0)
    : props.categories;
  const firstMatch = searching ? categories[0]?.tiles[0]?.id : undefined;
  const { onLaunch, onTileMenu, onCategoryMenu, onAddTile, onAddCategory, onMoveTile } = props;
  const [drag, setDrag] = useState<{ id: string; name: string; x: number; y: number } | null>(null);
  const [spot, setSpot] = useState<DropSpot | null>(null);
  /** 끌어 놓은 직후에 따라오는 클릭으로 타일이 열리지 않게 한다. */
  const justDragged = useRef(false);

  const beginDrag = (tile: Tile) => (e: ReactPointerEvent) => {
    if (e.button !== 0) return;
    const start = { x: e.clientX, y: e.clientY };
    let dragging = false;
    let latest: DropSpot | null = null;

    const move = (ev: PointerEvent) => {
      if (!dragging && Math.hypot(ev.clientX - start.x, ev.clientY - start.y) < DRAG_THRESHOLD) return;
      dragging = true;
      latest = spotAt(ev.clientX, ev.clientY);
      setSpot(latest);
      setDrag({ id: tile.id, name: tile.name, x: ev.clientX, y: ev.clientY });
    };
    const end = (ev: PointerEvent) => {
      window.removeEventListener("pointermove", move);
      window.removeEventListener("pointerup", end);
      window.removeEventListener("pointercancel", end);
      if (dragging) {
        justDragged.current = true;
        setTimeout(() => (justDragged.current = false));
        if (ev.type === "pointerup" && latest) onMoveTile(tile.id, latest.categoryId, latest.beforeTileId);
      }
      setDrag(null);
      setSpot(null);
    };
    window.addEventListener("pointermove", move);
    window.addEventListener("pointerup", end);
    window.addEventListener("pointercancel", end);
  };

  const launch = (tile: Tile) => {
    if (!justDragged.current) onLaunch(tile);
  };

  const menuAt = (e: MouseEvent) => {
    e.preventDefault();
    return { x: e.clientX, y: e.clientY };
  };

  return (
    <div className={`shortcuts tiles-${tileSize}`}>
      {categories.map((category) => (
        <section
          key={category.id}
          data-category-id={category.id}
          className={fileDropCategoryId === category.id ? "category file-drop" : "category"}
        >
          <header className="category-head" onContextMenu={(e) => onCategoryMenu(menuAt(e), category)}>
            <h3 title={category.name}>{category.name}</h3>
            <button className="icon-button" title="이 카테고리에 바로가기 추가" onClick={() => onAddTile(category.id)}>
              <PlusIcon />
            </button>
            <button
              className="icon-button"
              title="카테고리 메뉴"
              onClick={(e) => {
                const box = e.currentTarget.getBoundingClientRect();
                onCategoryMenu({ x: box.left, y: box.bottom + 2 }, category);
              }}
            >
              <MoreIcon />
            </button>
          </header>
          <div className={spot?.categoryId === category.id && spot.beforeTileId === null ? "tiles drop-end" : "tiles"}>
            {category.tiles.map((tile) => (
              <div
                key={tile.id}
                role="button"
                tabIndex={0}
                data-tile-id={tile.id}
                data-color={tile.color || undefined}
                title={tile.target}
                className={[
                  "tile",
                  drag?.id === tile.id ? "dragging" : "",
                  spot?.beforeTileId === tile.id && drag?.id !== tile.id ? "drop-before" : "",
                  tile.id === firstMatch ? "first-match" : "",
                ].join(" ")}
                onPointerDown={beginDrag(tile)}
                onClick={() => launch(tile)}
                onKeyDown={(e) => {
                  if (e.key === "Enter" || e.key === " ") {
                    e.preventDefault();
                    onLaunch(tile);
                  }
                }}
                onContextMenu={(e) => onTileMenu(menuAt(e), tile, category.id)}
              >
                <TileFace tile={tile} icon={iconFor(tile)} glyphSize={GLYPH_SIZE[tileSize]} />
                <span className="tile-name">{tile.name}</span>
              </div>
            ))}
            {category.tiles.length === 0 && !searching && (
              <button className="tiles-empty" onClick={() => onAddTile(category.id)}>
                <PlusIcon /> 바로가기 추가
              </button>
            )}
          </div>
        </section>
      ))}
      {searching && categories.length === 0 && <p className="no-match">‘{query.trim()}’ 와 맞는 바로가기가 없습니다</p>}
      {!searching && (
        <button className="add-category" onClick={onAddCategory}>
          <PlusIcon /> 카테고리 추가
        </button>
      )}
      {drag && (
        <div className="drag-ghost" style={{ left: drag.x + 12, top: drag.y + 12 }}>
          {drag.name}
        </div>
      )}
    </div>
  );
}
