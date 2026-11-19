import { useState } from "react";
import type { DragEvent, MouseEvent } from "react";
import type { Category, Tile, TileSize } from "../types";
import { KindIcon, MoreIcon, PlusIcon } from "./icons";

const ICON_SIZE: Record<TileSize, number> = { s: 16, m: 20, l: 26 };

interface Props {
  categories: Category[];
  tileSize: TileSize;
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

export function Shortcuts(props: Props) {
  const { categories, tileSize, onLaunch, onTileMenu, onCategoryMenu, onAddTile, onAddCategory, onMoveTile } = props;
  const [dragging, setDragging] = useState<string | null>(null);
  const [spot, setSpot] = useState<DropSpot | null>(null);

  const endDrag = () => {
    setDragging(null);
    setSpot(null);
  };

  const hover = (e: DragEvent, next: DropSpot) => {
    if (!dragging) return;
    e.preventDefault();
    e.stopPropagation();
    e.dataTransfer.dropEffect = "move";
    if (spot?.categoryId !== next.categoryId || spot.beforeTileId !== next.beforeTileId) setSpot(next);
  };

  const drop = (e: DragEvent) => {
    if (!dragging) return;
    e.preventDefault();
    e.stopPropagation();
    if (spot) onMoveTile(dragging, spot.categoryId, spot.beforeTileId);
    endDrag();
  };

  const menuAt = (e: MouseEvent) => {
    e.preventDefault();
    return { x: e.clientX, y: e.clientY };
  };

  return (
    <div className={`shortcuts tiles-${tileSize}`}>
      {categories.map((category) => (
        <section key={category.id} className="category">
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
          <div
            className={spot?.categoryId === category.id && spot.beforeTileId === null ? "tiles drop-end" : "tiles"}
            onDragOver={(e) => hover(e, { categoryId: category.id, beforeTileId: null })}
            onDrop={drop}
          >
            {category.tiles.map((tile, index) => (
              <div
                key={tile.id}
                role="button"
                tabIndex={0}
                draggable
                title={tile.target}
                className={[
                  "tile",
                  dragging === tile.id ? "dragging" : "",
                  spot?.beforeTileId === tile.id && dragging !== tile.id ? "drop-before" : "",
                ].join(" ")}
                onClick={() => onLaunch(tile)}
                onKeyDown={(e) => {
                  if (e.key === "Enter" || e.key === " ") {
                    e.preventDefault();
                    onLaunch(tile);
                  }
                }}
                onContextMenu={(e) => onTileMenu(menuAt(e), tile, category.id)}
                onDragStart={(e) => {
                  e.dataTransfer.effectAllowed = "move";
                  e.dataTransfer.setData("text/plain", tile.name);
                  setDragging(tile.id);
                }}
                onDragEnd={endDrag}
                onDragOver={(e) => {
                  // 타일의 왼쪽 절반에 놓으면 그 앞, 오른쪽 절반이면 그 뒤에 들어간다.
                  const box = e.currentTarget.getBoundingClientRect();
                  const after = e.clientX > box.left + box.width / 2;
                  const beforeTileId = after ? (category.tiles[index + 1]?.id ?? null) : tile.id;
                  hover(e, { categoryId: category.id, beforeTileId });
                }}
                onDrop={drop}
              >
                <span className={`tile-icon kind-${tile.kind}`}>
                  <KindIcon kind={tile.kind} size={ICON_SIZE[tileSize]} />
                </span>
                <span className="tile-name">{tile.name}</span>
              </div>
            ))}
            {category.tiles.length === 0 && (
              <button className="tiles-empty" onClick={() => onAddTile(category.id)}>
                <PlusIcon /> 바로가기 추가
              </button>
            )}
          </div>
        </section>
      ))}
      <button className="add-category" onClick={onAddCategory}>
        <PlusIcon /> 카테고리 추가
      </button>
    </div>
  );
}
