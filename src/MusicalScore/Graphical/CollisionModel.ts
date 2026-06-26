import { RectangleF2D } from "../../Common/DataObjects/RectangleF2D";
import { BoundingBox } from "./BoundingBox";
import type { GraphicalMusicSheet } from "./GraphicalMusicSheet";

export enum CollisionBoxKind {
    Articulation = "Articulation",
    Beam = "Beam",
    GenericBoundingBox = "BoundingBox",
    Fingering = "Fingering",
    MeasureBarline = "MeasureBarline",
    Notehead = "Notehead",
    Ornament = "Ornament",
    Stem = "Stem",
    Tie = "Tie",
    Tuplet = "Tuplet",
    Unknown = "Unknown",
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

    public clear(): void {
        this.boxes = [];
        this.nextId = 1;
    }

    public get Boxes(): CollisionBox[] {
        return this.boxes;
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
        const className: string = (boundingBox.DataObject as any)?.constructor?.name ?? "";
        if (className.indexOf("Label") >= 0) {
            const labelText: string = (boundingBox.DataObject as any)?.Label?.text;
            if (/^\s*\d+\s*$/.test(labelText ?? "")) {
                return CollisionBoxKind.Fingering;
            }
            return CollisionBoxKind.GenericBoundingBox;
        }
        if (className.indexOf("Note") >= 0) {
            return CollisionBoxKind.Notehead;
        }
        if (className.indexOf("Slur") >= 0 || className.indexOf("Tie") >= 0) {
            return CollisionBoxKind.Tie;
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
