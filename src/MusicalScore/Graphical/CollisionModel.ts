import { RectangleF2D } from "../../Common/DataObjects/RectangleF2D";
import { BoundingBox } from "./BoundingBox";
import type { GraphicalMusicSheet } from "./GraphicalMusicSheet";

export enum CollisionBoxKind {
    Accidental = "Accidental",
    Articulation = "Articulation",
    Beam = "Beam",
    ChordSymbol = "ChordSymbol",
    Clef = "Clef",
    Connector = "Connector",
    Dot = "Dot",
    Dynamic = "Dynamic",
    Expression = "Expression",
    Flag = "Flag",
    GenericBoundingBox = "BoundingBox",
    Fingering = "Fingering",
    Glissando = "Glissando",
    InstrumentLabel = "InstrumentLabel",
    KeySignature = "KeySignature",
    Label = "Label",
    LedgerLine = "LedgerLine",
    Lyric = "Lyric",
    MeasureBarline = "MeasureBarline",
    MeasureNumber = "MeasureNumber",
    Note = "Note",
    Notehead = "Notehead",
    Ornament = "Ornament",
    OctaveShift = "OctaveShift",
    Pedal = "Pedal",
    Repeat = "Repeat",
    RepeatDot = "RepeatDot",
    Rest = "Rest",
    Slur = "Slur",
    StaffLine = "StaffLine",
    Stem = "Stem",
    SystemLine = "SystemLine",
    Text = "Text",
    Tie = "Tie",
    TimeSignature = "TimeSignature",
    Tuplet = "Tuplet",
    Unknown = "Unknown",
    Volta = "Volta",
    WavyLine = "WavyLine",
}

export interface CollisionRect {
    x: number;
    y: number;
    width: number;
    height: number;
}

export interface CollisionBox {
    id: number;
    rect: CollisionRect;
    kind: CollisionBoxKind;
    owner?: Object;
    source?: Object;
    movable: boolean;
    priority: number;
    label?: string;
}

export interface CollisionQueryOptions {
    includeMargins?: boolean;
    padding?: number;
    ignoreIds?: Set<number>;
    ignoreOwners?: Set<Object>;
    ignoreKinds?: Set<CollisionBoxKind>;
}

export interface CollisionMagnetOptions extends CollisionQueryOptions {
    stepX?: number;
    stepY?: number;
    maxDistance?: number;
    preferredDirections?: CollisionMagnetDirection[];
}

export enum CollisionMagnetDirection {
    Up = "Up",
    Down = "Down",
    Left = "Left",
    Right = "Right",
    Stay = "Stay",
}

interface CollisionCandidate {
    rect: CollisionRect;
    distance: number;
    directionPenalty: number;
}

export class CollisionModel {
    private static readonly containerClassNames: Set<string> = new Set<string>([
        "GraphicalMusicPage",
        "MusicSystem",
        "VexFlowMusicSystem",
        "StaffLine",
        "VexFlowStaffLine",
        "GraphicalMeasure",
        "VexFlowMeasure",
        "VexFlowMultiRestMeasure",
        "VexFlowTabMeasure",
        "GraphicalStaffEntry",
        "VexFlowStaffEntry",
        "GraphicalVoiceEntry",
        "VexFlowVoiceEntry",
    ]);

    private boxes: CollisionBox[] = [];
    private nextId: number = 1;
    private kindIndexes: Map<string, CollisionKindSpatialIndex> = new Map<string, CollisionKindSpatialIndex>();

    public clear(): void {
        this.boxes = [];
        this.nextId = 1;
        this.kindIndexes.clear();
    }

    public get Boxes(): CollisionBox[] {
        return this.boxes;
    }

    /**
     * Returns, in registration order, every box of one of the given kinds whose rect could overlap the
     * query region (in units). The result is a conservative superset: callers still run their exact
     * overlap test on each returned box. Boxes registered since the last query are indexed lazily.
     * @param indexKey identifies the kind set, so each distinct set keeps its own index.
     */
    public queryBoxesOfKinds(indexKey: string, kinds: ReadonlySet<CollisionBoxKind>,
                             left: number, top: number, right: number, bottom: number): CollisionBox[] {
        let index: CollisionKindSpatialIndex = this.kindIndexes.get(indexKey);
        if (!index || index.boxes !== this.boxes) {
            index = new CollisionKindSpatialIndex(this.boxes, kinds);
            this.kindIndexes.set(indexKey, index);
        }
        return index.query(left, top, right, bottom);
    }

    public registerBoundingBoxesFromMusicSheet(graphicalMusicSheet: GraphicalMusicSheet): void {
        for (const page of graphicalMusicSheet.MusicPages) {
            this.registerBoundingBoxTree(page.PositionAndShape);
        }
    }

    public registerBoundingBoxTree(root: BoundingBox): void {
        this.registerBoundingBoxTreeRecursive(root);
    }

    public registerBoundingBox(boundingBox: BoundingBox, kind: CollisionBoxKind = CollisionBoxKind.GenericBoundingBox): CollisionBox {
        const rect: CollisionRect = CollisionModel.rectFromBoundingBox(boundingBox);
        if (!CollisionModel.isUsableRect(rect)) {
            return undefined;
        }
        return this.registerRect(rect, kind, boundingBox.DataObject, boundingBox);
    }

    public registerRect(
        rect: CollisionRect,
        kind: CollisionBoxKind = CollisionBoxKind.Unknown,
        owner?: Object,
        source?: Object,
        movable: boolean = false,
        priority: number = 0,
        label?: string
    ): CollisionBox {
        const normalizedRect: CollisionRect = CollisionModel.normalizeRect(rect);
        if (!CollisionModel.isUsableRect(normalizedRect)) {
            return undefined;
        }
        const collisionBox: CollisionBox = {
            id: this.nextId++,
            rect: normalizedRect,
            kind,
            owner,
            source,
            movable,
            priority,
            label,
        };
        this.boxes.push(collisionBox);
        return collisionBox;
    }

    public collides(rect: CollisionRect, options: CollisionQueryOptions = {}): boolean {
        return this.getCollisions(rect, options).length > 0;
    }

    public getCollisions(rect: CollisionRect, options: CollisionQueryOptions = {}): CollisionBox[] {
        const queryRect: CollisionRect = CollisionModel.inflateRect(CollisionModel.normalizeRect(rect), options.padding ?? 0);
        if (!CollisionModel.isUsableRect(queryRect)) {
            return [];
        }
        return this.boxes.filter((box: CollisionBox) => {
            if (options.ignoreIds?.has(box.id) ||
                options.ignoreOwners?.has(box.owner) ||
                options.ignoreKinds?.has(box.kind)) {
                return false;
            }
            return CollisionModel.rectsOverlap(queryRect, box.rect);
        });
    }

    public findClosestNonCollidingRect(rect: CollisionRect, options: CollisionMagnetOptions = {}): CollisionRect {
        const baseRect: CollisionRect = CollisionModel.normalizeRect(rect);
        if (!this.collides(baseRect, options)) {
            return baseRect;
        }

        const stepX: number = Math.max(options.stepX ?? 0.1, 0.01);
        const stepY: number = Math.max(options.stepY ?? 0.1, 0.01);
        const maxDistance: number = options.maxDistance ?? 3;
        const preferredDirections: CollisionMagnetDirection[] =
            options.preferredDirections ?? [
                CollisionMagnetDirection.Stay,
                CollisionMagnetDirection.Up,
                CollisionMagnetDirection.Down,
                CollisionMagnetDirection.Left,
                CollisionMagnetDirection.Right,
            ];
        const candidates: CollisionCandidate[] = [];

        for (const direction of preferredDirections) {
            const directionPenalty: number = preferredDirections.indexOf(direction) * 0.001;
            switch (direction) {
                case CollisionMagnetDirection.Stay:
                    candidates.push(...this.generateGridCandidates(baseRect, stepX, stepY, maxDistance, directionPenalty));
                    break;
                case CollisionMagnetDirection.Up:
                    candidates.push(...this.generateAxisCandidates(baseRect, 0, -stepY, maxDistance, directionPenalty));
                    break;
                case CollisionMagnetDirection.Down:
                    candidates.push(...this.generateAxisCandidates(baseRect, 0, stepY, maxDistance, directionPenalty));
                    break;
                case CollisionMagnetDirection.Left:
                    candidates.push(...this.generateAxisCandidates(baseRect, -stepX, 0, maxDistance, directionPenalty));
                    break;
                case CollisionMagnetDirection.Right:
                    candidates.push(...this.generateAxisCandidates(baseRect, stepX, 0, maxDistance, directionPenalty));
                    break;
                default:
                    break;
            }
        }

        candidates.sort((a: CollisionCandidate, b: CollisionCandidate) =>
            (a.distance + a.directionPenalty) - (b.distance + b.directionPenalty));

        for (const candidate of candidates) {
            if (!this.collides(candidate.rect, options)) {
                return candidate.rect;
            }
        }
        return baseRect;
    }

    public static rectFromBoundingBox(boundingBox: BoundingBox): CollisionRect {
        return {
            x: boundingBox.AbsolutePosition.x + boundingBox.BorderLeft,
            y: boundingBox.AbsolutePosition.y + boundingBox.BorderTop,
            width: boundingBox.BorderRight - boundingBox.BorderLeft,
            height: boundingBox.BorderBottom - boundingBox.BorderTop,
        };
    }

    public static rectsOverlap(a: CollisionRect, b: CollisionRect): boolean {
        return Math.min(a.x + a.width, b.x + b.width) > Math.max(a.x, b.x) &&
            Math.min(a.y + a.height, b.y + b.height) > Math.max(a.y, b.y);
    }

    public static inflateRect(rect: CollisionRect, padding: number): CollisionRect {
        if (padding === 0) {
            return rect;
        }
        return {
            x: rect.x - padding,
            y: rect.y - padding,
            width: rect.width + padding * 2,
            height: rect.height + padding * 2,
        };
    }

    public static normalizeRect(rect: CollisionRect): CollisionRect {
        const x2: number = rect.x + rect.width;
        const y2: number = rect.y + rect.height;
        return {
            x: Math.min(rect.x, x2),
            y: Math.min(rect.y, y2),
            width: Math.abs(rect.width),
            height: Math.abs(rect.height),
        };
    }

    public static isUsableRect(rect: CollisionRect): boolean {
        return rect !== undefined &&
            Number.isFinite(rect.x) &&
            Number.isFinite(rect.y) &&
            Number.isFinite(rect.width) &&
            Number.isFinite(rect.height) &&
            rect.width > 0.0001 &&
            rect.height > 0.0001;
    }

    public static toRectangleF2D(rect: CollisionRect): RectangleF2D {
        return new RectangleF2D(rect.x, rect.y, rect.width, rect.height);
    }

    private registerBoundingBoxTreeRecursive(boundingBox: BoundingBox): void {
        const dataObject: any = boundingBox.DataObject;
        const className: string = dataObject?.constructor?.name;
        const isContainer: boolean = CollisionModel.containerClassNames.has(className);
        const isLeaf: boolean = boundingBox.ChildElements.length === 0;
        if (!isContainer && isLeaf) {
            this.registerBoundingBox(boundingBox, this.classifyBoundingBox(boundingBox));
        }
        for (const child of boundingBox.ChildElements) {
            this.registerBoundingBoxTreeRecursive(child);
        }
    }

    private classifyBoundingBox(boundingBox: BoundingBox): CollisionBoxKind {
        const dataObject: any = boundingBox.DataObject as any;
        const className: string = dataObject?.constructor?.name ?? "";
        const parentClassName: string = (boundingBox.Parent?.DataObject as any)?.constructor?.name ?? "";
        if (className.indexOf("Label") >= 0) {
            const labelText: string = dataObject?.Label?.text;
            if (/^\s*\d+\s*$/.test(labelText ?? "")) {
                return parentClassName.indexOf("StaffLine") >= 0
                    ? CollisionBoxKind.Fingering
                    : CollisionBoxKind.MeasureNumber;
            }
            if (parentClassName.indexOf("Lyric") >= 0) {
                return CollisionBoxKind.Lyric;
            }
            if (parentClassName.indexOf("ChordSymbol") >= 0) {
                return CollisionBoxKind.ChordSymbol;
            }
            if (parentClassName.indexOf("MusicSystem") >= 0) {
                return CollisionBoxKind.InstrumentLabel;
            }
            return CollisionBoxKind.Label;
        }
        if (className.indexOf("Note") >= 0) {
            if (typeof dataObject?.sourceNote?.isRest === "function" && dataObject.sourceNote.isRest()) {
                return CollisionBoxKind.Rest;
            }
            return CollisionBoxKind.Notehead;
        }
        if (className.indexOf("Slur") >= 0) {
            return CollisionBoxKind.Slur;
        }
        if (className.indexOf("Tie") >= 0) {
            return CollisionBoxKind.Tie;
        }
        if (className.indexOf("Glissando") >= 0) {
            return CollisionBoxKind.Glissando;
        }
        if (className.indexOf("OctaveShift") >= 0) {
            return CollisionBoxKind.OctaveShift;
        }
        if (className.indexOf("Pedal") >= 0) {
            return CollisionBoxKind.Pedal;
        }
        if (className.indexOf("WavyLine") >= 0 || className.indexOf("Vibrato") >= 0) {
            return CollisionBoxKind.WavyLine;
        }
        if (className.indexOf("DynamicExpression") >= 0) {
            return CollisionBoxKind.Dynamic;
        }
        if (className.indexOf("Expression") >= 0) {
            return CollisionBoxKind.Expression;
        }
        if (className.indexOf("SystemLine") >= 0) {
            return CollisionBoxKind.SystemLine;
        }
        if (className.indexOf("InstrumentBracket") >= 0 || className.indexOf("InstrumentBrace") >= 0) {
            return CollisionBoxKind.Connector;
        }
        return CollisionBoxKind.GenericBoundingBox;
    }

    private generateAxisCandidates(
        baseRect: CollisionRect,
        stepX: number,
        stepY: number,
        maxDistance: number,
        directionPenalty: number
    ): CollisionCandidate[] {
        const candidates: CollisionCandidate[] = [];
        const stepDistance: number = Math.max(Math.abs(stepX), Math.abs(stepY), 0.01);
        for (let distance: number = stepDistance; distance <= maxDistance + 0.0001; distance += stepDistance) {
            const multiplier: number = distance / stepDistance;
            candidates.push({
                rect: {
                    x: baseRect.x + stepX * multiplier,
                    y: baseRect.y + stepY * multiplier,
                    width: baseRect.width,
                    height: baseRect.height,
                },
                distance,
                directionPenalty,
            });
        }
        return candidates;
    }

    private generateGridCandidates(
        baseRect: CollisionRect,
        stepX: number,
        stepY: number,
        maxDistance: number,
        directionPenalty: number
    ): CollisionCandidate[] {
        const candidates: CollisionCandidate[] = [];
        const xSteps: number = Math.ceil(maxDistance / stepX);
        const ySteps: number = Math.ceil(maxDistance / stepY);
        for (let ix: number = -xSteps; ix <= xSteps; ix++) {
            for (let iy: number = -ySteps; iy <= ySteps; iy++) {
                if (ix === 0 && iy === 0) {
                    continue;
                }
                const dx: number = ix * stepX;
                const dy: number = iy * stepY;
                const distance: number = Math.sqrt(dx * dx + dy * dy);
                if (distance > maxDistance) {
                    continue;
                }
                candidates.push({
                    rect: {
                        x: baseRect.x + dx,
                        y: baseRect.y + dy,
                        width: baseRect.width,
                        height: baseRect.height,
                    },
                    distance,
                    directionPenalty,
                });
            }
        }
        return candidates;
    }
}

/**
 * Uniform-grid index over the boxes of a CollisionModel that have one of a fixed set of kinds.
 * Only used to prefilter candidates; it never decides a collision by itself.
 */
class CollisionKindSpatialIndex {
    private static readonly cellSize: number = 4;
    private static readonly maxCellsPerBox: number = 256;
    private static readonly cellRowSpan: number = 2097152;
    private static readonly cellRowOffset: number = 1048576;

    public readonly boxes: CollisionBox[];
    private readonly kinds: ReadonlySet<CollisionBoxKind>;
    private readonly cells: Map<number, number[]> = new Map<number, number[]>();
    /** boxes too large (or too far out) for the grid; returned by every query */
    private readonly unbucketed: number[] = [];
    private readonly seenStamps: number[] = [];
    private indexedCount: number = 0;
    private queryStamp: number = 0;

    constructor(boxes: CollisionBox[], kinds: ReadonlySet<CollisionBoxKind>) {
        this.boxes = boxes;
        this.kinds = kinds;
    }

    public query(left: number, top: number, right: number, bottom: number): CollisionBox[] {
        this.indexNewBoxes();
        const result: CollisionBox[] = [];
        const cellSize: number = CollisionKindSpatialIndex.cellSize;
        const cx0: number = Math.floor(left / cellSize) - 1;
        const cx1: number = Math.floor(right / cellSize) + 1;
        const cy0: number = Math.floor(top / cellSize) - 1;
        const cy1: number = Math.floor(bottom / cellSize) + 1;
        const indices: number[] = this.unbucketed.slice();
        if (!this.isGridRange(cx0, cx1, cy0, cy1) ||
            (cx1 - cx0 + 1) * (cy1 - cy0 + 1) > CollisionKindSpatialIndex.maxCellsPerBox * 16) {
            // query region too large for the grid: fall back to every indexed box
            for (let i: number = 0; i < this.indexedCount; i++) {
                if (this.kinds.has(this.boxes[i].kind)) {
                    result.push(this.boxes[i]);
                }
            }
            return result;
        }
        const stamp: number = ++this.queryStamp;
        for (const index of indices) {
            this.seenStamps[index] = stamp;
        }
        for (let cx: number = cx0; cx <= cx1; cx++) {
            for (let cy: number = cy0; cy <= cy1; cy++) {
                const cell: number[] = this.cells.get(this.cellKey(cx, cy));
                if (!cell) {
                    continue;
                }
                for (const index of cell) {
                    if (this.seenStamps[index] !== stamp) {
                        this.seenStamps[index] = stamp;
                        indices.push(index);
                    }
                }
            }
        }
        indices.sort((a: number, b: number) => a - b);
        for (const index of indices) {
            result.push(this.boxes[index]);
        }
        return result;
    }

    private indexNewBoxes(): void {
        const cellSize: number = CollisionKindSpatialIndex.cellSize;
        for (; this.indexedCount < this.boxes.length; this.indexedCount++) {
            const index: number = this.indexedCount;
            const box: CollisionBox = this.boxes[index];
            if (!box || !this.kinds.has(box.kind)) {
                continue;
            }
            const rect: CollisionRect = box.rect;
            const cx0: number = Math.floor(rect.x / cellSize);
            const cx1: number = Math.floor((rect.x + rect.width) / cellSize);
            const cy0: number = Math.floor(rect.y / cellSize);
            const cy1: number = Math.floor((rect.y + rect.height) / cellSize);
            if (!this.isGridRange(cx0, cx1, cy0, cy1) ||
                (cx1 - cx0 + 1) * (cy1 - cy0 + 1) > CollisionKindSpatialIndex.maxCellsPerBox) {
                this.unbucketed.push(index);
                continue;
            }
            for (let cx: number = cx0; cx <= cx1; cx++) {
                for (let cy: number = cy0; cy <= cy1; cy++) {
                    const key: number = this.cellKey(cx, cy);
                    let cell: number[] = this.cells.get(key);
                    if (!cell) {
                        cell = [];
                        this.cells.set(key, cell);
                    }
                    cell.push(index);
                }
            }
        }
    }

    private isGridRange(cx0: number, cx1: number, cy0: number, cy1: number): boolean {
        const limit: number = CollisionKindSpatialIndex.cellRowOffset;
        return Number.isFinite(cx0) && Number.isFinite(cx1) && Number.isFinite(cy0) && Number.isFinite(cy1) &&
            cx0 > -limit && cx1 < limit && cy0 > -limit && cy1 < limit;
    }

    private cellKey(cx: number, cy: number): number {
        return cx * CollisionKindSpatialIndex.cellRowSpan + (cy + CollisionKindSpatialIndex.cellRowOffset);
    }
}
