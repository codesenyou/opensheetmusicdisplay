import Vex from "vexflow";
import VF = Vex.Flow;
import { GraphicalStaffEntry } from "./GraphicalStaffEntry";
import { StaffLine } from "./StaffLine";
import { GraphicalMusicSheet } from "./GraphicalMusicSheet";
import { EngravingRules } from "./EngravingRules";
import { Tie } from "../VoiceData/Tie";
import { Fraction } from "../../Common/DataObjects/Fraction";
import { Note } from "../VoiceData/Note";
import { MusicSheet } from "../MusicSheet";
import { GraphicalMeasure } from "./GraphicalMeasure";
import {ClefInstruction, ClefEnum} from "../VoiceData/Instructions/ClefInstruction";
import { LyricWord } from "../VoiceData/Lyrics/LyricsWord";
import { SourceMeasure } from "../VoiceData/SourceMeasure";
import { GraphicalMusicPage } from "./GraphicalMusicPage";
import { GraphicalNote } from "./GraphicalNote";
import { Beam } from "../VoiceData/Beam";
import { OctaveEnum } from "../VoiceData/Expressions/ContinuousExpressions/OctaveShift";
import { VoiceEntry, StemDirectionType } from "../VoiceData/VoiceEntry";
import { OrnamentContainer, OrnamentEnum } from "../VoiceData/OrnamentContainer";
import { Articulation } from "../VoiceData/Articulation";
import { Tuplet } from "../VoiceData/Tuplet";
import { MusicSystem } from "./MusicSystem";
import { GraphicalTie } from "./GraphicalTie";
import { RepetitionInstruction, RepetitionInstructionEnum, AlignmentType } from "../VoiceData/Instructions/RepetitionInstruction";
import { MultiExpression, MultiExpressionEntry } from "../VoiceData/Expressions/MultiExpression";
import { StaffEntryLink } from "../VoiceData/StaffEntryLink";
import { MusicSystemBuilder } from "./MusicSystemBuilder";
import { MultiTempoExpression } from "../VoiceData/Expressions/MultiTempoExpression";
import { Repetition } from "../MusicSource/Repetition";
import { PointF2D } from "../../Common/DataObjects/PointF2D";
import { SourceStaffEntry } from "../VoiceData/SourceStaffEntry";
import { BoundingBox } from "./BoundingBox";
import { Instrument } from "../Instrument";
import { GraphicalLabel } from "./GraphicalLabel";
import { TextAlignmentEnum } from "../../Common/Enums/TextAlignment";
import { VerticalGraphicalStaffEntryContainer } from "./VerticalGraphicalStaffEntryContainer";
import { KeyInstruction } from "../VoiceData/Instructions/KeyInstruction";
import { AbstractNotationInstruction } from "../VoiceData/Instructions/AbstractNotationInstruction";
import { TechnicalInstruction, TechnicalInstructionType } from "../VoiceData/Instructions/TechnicalInstruction";
import { AccidentalEnum, Pitch } from "../../Common/DataObjects/Pitch";
import { LinkedVoice } from "../VoiceData/LinkedVoice";
import { IGraphicalSymbolFactory } from "../Interfaces/IGraphicalSymbolFactory";
import { ITextMeasurer } from "../Interfaces/ITextMeasurer";
import { ITransposeCalculator } from "../Interfaces/ITransposeCalculator";
import { OctaveShiftParams } from "./OctaveShiftParams";
import { AccidentalCalculator } from "./AccidentalCalculator";
import { MidiInstrument } from "../VoiceData/Instructions/ClefInstruction";
import { Staff } from "../VoiceData/Staff";
import { OctaveShift } from "../VoiceData/Expressions/ContinuousExpressions/OctaveShift";
import { NoteHeadShape } from "../VoiceData/Notehead";
import log from "loglevel";
import { Dictionary } from "typescript-collections";
import { GraphicalLyricEntry } from "./GraphicalLyricEntry";
import { GraphicalLyricWord } from "./GraphicalLyricWord";
import { GraphicalLine } from "./GraphicalLine";
import { Label } from "../Label";
import { GraphicalVoiceEntry } from "./GraphicalVoiceEntry";
import { VerticalSourceStaffEntryContainer } from "../VoiceData/VerticalSourceStaffEntryContainer";
import { SkyBottomLineCalculator } from "./SkyBottomLineCalculator";
import { PlacementEnum } from "../VoiceData/Expressions/AbstractExpression";
import { AbstractGraphicalInstruction } from "./AbstractGraphicalInstruction";
import { GraphicalInstantaneousTempoExpression } from "./GraphicalInstantaneousTempoExpression";
import { InstantaneousTempoExpression, TempoType } from "../VoiceData/Expressions/InstantaneousTempoExpression";
import { ContinuousTempoExpression } from "../VoiceData/Expressions/ContinuousExpressions/ContinuousTempoExpression";
import { FontStyles } from "../../Common/Enums/FontStyles";
import { AbstractTempoExpression } from "../VoiceData/Expressions/AbstractTempoExpression";
import { GraphicalInstantaneousDynamicExpression } from "./GraphicalInstantaneousDynamicExpression";
import { ContDynamicEnum } from "../VoiceData/Expressions/ContinuousExpressions/ContinuousDynamicExpression";
import { GraphicalContinuousDynamicExpression } from "./GraphicalContinuousDynamicExpression";
import { FillEmptyMeasuresWithWholeRests } from "../../OpenSheetMusicDisplay/OSMDOptions";
import { IStafflineNoteCalculator } from "../Interfaces/IStafflineNoteCalculator";
import { GraphicalUnknownExpression } from "./GraphicalUnknownExpression";
import { GraphicalChordSymbolContainer } from "./GraphicalChordSymbolContainer";
import { LyricsEntry } from "../VoiceData/Lyrics/LyricsEntry";
import { Voice } from "../VoiceData/Voice";
import { TabNote } from "../VoiceData/TabNote";

interface FingeringCollisionRect {
    left: number;
    right: number;
    top: number;
    bottom: number;
    penalty: number;
    maxShift?: number;
    soft?: boolean;
    isFingering?: boolean;
    kind?: string;
    debugMeasure?: number;
    debugMeasureX?: number;
    debugStaffId?: number;
    debugKey?: string;
}

interface FingeringPxRect {
    x: number;
    y: number;
    width: number;
    height: number;
}

interface FingeringOrnamentVisualMetrics {
    width: number;
    height: number;
    bottomOffset: number;
    advanceHeight: number;
}

interface FingeringPlacementWorkItem {
    fingering: TechnicalInstruction;
    label: GraphicalLabel;
    placement: PlacementEnum;
    stackIndex: number;
    ownerCenterX: number;
    initialY: number;
    labelHeight: number;
    initialRect: FingeringCollisionRect;
    placedRect?: FingeringCollisionRect;
}

interface FingeringLabelCandidate {
    dx: number;
    dy: number;
    rect: FingeringCollisionRect;
    cost: number;
}

interface FingeringGroupResolution {
    dy: number;
    candidates: FingeringLabelCandidate[];
    cost: number;
}

interface FingeringDebugSvgRect {
    x: number;
    y: number;
    width: number;
    height: number;
}

interface FingeringDebugRect extends FingeringCollisionRect {
    index?: number;
    svgRect?: FingeringDebugSvgRect;
    overlapsInitialLabel?: boolean;
    overlapsPlacedLabel?: boolean;
}

type InternalCollisionDebugKind =
    "beam" | "stem" | "note" | "notehead" | "barline" | "existing-fingering" |
    "label-initial" | "label-placed" | "tie" | "ornament" | "articulation" | "ornament-reserve" | "rest";

interface InternalCollisionDebugBox {
    key: string;
    kind: InternalCollisionDebugKind;
    measure: number;
    staffId: number;
    measureX: number;
    left: number;
    right: number;
    top: number;
    bottom: number;
    source: "fingering-calculator";
    collision?: boolean;
    reason?: string;
    fingering?: string;
    placement?: string;
    stackIndex?: number;
    svgRect?: FingeringDebugSvgRect;
}

/**
 * Class used to do all the calculations in a MusicSheet, which in the end populates a GraphicalMusicSheet.
 */
export abstract class MusicSheetCalculator {
    public static symbolFactory: IGraphicalSymbolFactory;
    public static transposeCalculator: ITransposeCalculator;
    public static stafflineNoteCalculator: IStafflineNoteCalculator;
    protected static textMeasurer: ITextMeasurer;

    protected staffEntriesWithGraphicalTies: GraphicalStaffEntry[] = [];
    protected staffEntriesWithOrnaments: GraphicalStaffEntry[] = [];
    protected staffEntriesWithChordSymbols: GraphicalStaffEntry[] = [];
    protected staffLinesWithLyricWords: StaffLine[] = [];

    protected graphicalLyricWords: GraphicalLyricWord[] = [];

    protected graphicalMusicSheet: GraphicalMusicSheet;
    protected rules: EngravingRules;
    protected musicSystems: MusicSystem[];

    private abstractNotImplementedErrorMessage: string = "abstract, not implemented";
    private fingeringInternalDebugBoxesByKey: Map<string, InternalCollisionDebugBox> =
        new Map<string, InternalCollisionDebugBox>();

    public static get TextMeasurer(): ITextMeasurer {
        return MusicSheetCalculator.textMeasurer;
    }

    public static set TextMeasurer(value: ITextMeasurer) {
        MusicSheetCalculator.textMeasurer = value;
    }

    protected get leadSheet(): boolean {
        return this.graphicalMusicSheet.LeadSheet;
    }

    protected static setMeasuresMinStaffEntriesWidth(measures: GraphicalMeasure[], minimumStaffEntriesWidth: number): void {
        for (let idx: number = 0, len: number = measures.length; idx < len; ++idx) {
            const measure: GraphicalMeasure = measures[idx];
            if (measure) {
                measure.minimumStaffEntriesWidth = minimumStaffEntriesWidth;
            }
        }
    }

    public initialize(graphicalMusicSheet: GraphicalMusicSheet): void {
        this.graphicalMusicSheet = graphicalMusicSheet;
        this.rules = graphicalMusicSheet.ParentMusicSheet.Rules;
        this.rules.clearMusicSheetObjects();
        this.prepareGraphicalMusicSheet();
        //this.calculate();
    }

    /**
     * Build the 2D [[GraphicalMeasure]] list needed for the [[MusicSheetCalculator]].
     * Internally it creates [[GraphicalMeasure]]s, [[GraphicalStaffEntry]]'s and [[GraphicalNote]]s.
     */
    public prepareGraphicalMusicSheet(): void {
        // Clear the stored system images dict - all systems have to be redrawn.
        // Not necessary now. TODO Check
        // this.graphicalMusicSheet.SystemImages.length = 0;
        const musicSheet: MusicSheet = this.graphicalMusicSheet.ParentMusicSheet;

        this.staffEntriesWithGraphicalTies = [];
        this.staffEntriesWithOrnaments = [];
        this.staffEntriesWithChordSymbols = [];
        this.staffLinesWithLyricWords = [];
        // this.staffLinesWithGraphicalExpressions = [];

        this.graphicalMusicSheet.Initialize();
        const measureList: GraphicalMeasure[][] = this.graphicalMusicSheet.MeasureList;

        // one AccidentalCalculator for each Staff (regardless of Instrument)
        const accidentalCalculators: AccidentalCalculator[] = this.createAccidentalCalculators();

        // List of Active ClefInstructions
        const activeClefs: ClefInstruction[] = this.graphicalMusicSheet.initializeActiveClefs();

        // LyricWord - GraphicalLyricWord Lists
        const lyricWords: LyricWord[] = [];

        const completeNumberOfStaves: number = musicSheet.getCompleteNumberOfStaves();

        // Octave Shifts List
        const openOctaveShifts: OctaveShiftParams[] = [];

        // TieList - timestampsArray
        for (let i: number = 0; i < completeNumberOfStaves; i++) {
            openOctaveShifts.push(undefined);
        }

        // go through all SourceMeasures (taking into account normal SourceMusicParts and Repetitions)
        for (let idx: number = 0, len: number = musicSheet.SourceMeasures.length; idx < len; ++idx) {
            const sourceMeasure: SourceMeasure = musicSheet.SourceMeasures[idx];
            const graphicalMeasures: GraphicalMeasure[] = this.createGraphicalMeasuresForSourceMeasure(
                sourceMeasure,
                accidentalCalculators,
                lyricWords,
                openOctaveShifts,
                activeClefs
            );
            measureList.push(graphicalMeasures);
            if (sourceMeasure.multipleRestMeasures > 0 && this.rules.RenderMultipleRestMeasures) {
                // multiRest given in XML, skip the next measures included
                sourceMeasure.isReducedToMultiRest = true;
                sourceMeasure.multipleRestMeasureNumber = 1;
                const measuresToSkip: number = sourceMeasure.multipleRestMeasures - 1;
                // console.log(`skipping ${measuresToSkip} measures for measure #${sourceMeasure.MeasureNumber}.`);
                idx += measuresToSkip;
                for (let idx2: number = 1; idx2 <= measuresToSkip; idx2++) {
                    const nextMeasureIndex: number = musicSheet.SourceMeasures.indexOf(sourceMeasure) + idx2;
                    // note that if there are pickup measures in the sheet, the measure index is not MeasureNumber - 1.
                    //   (if first measure in the sheet is a pickup measure, its index and measure number will be 0)
                    if (nextMeasureIndex >= musicSheet.SourceMeasures.length) {
                        break; // shouldn't happen, but for safety.
                    }
                    const nextSourceMeasure: SourceMeasure = musicSheet.SourceMeasures[nextMeasureIndex];
                    // TODO handle the case that a measure after the first multiple rest measure can't be reduced
                    nextSourceMeasure.multipleRestMeasureNumber = idx2 + 1;
                    nextSourceMeasure.isReducedToMultiRest = true;
                    measureList.push([undefined]);
                    // TODO we could push an object here or push nothing entirely,
                    //   but then the index doesn't correspond to measure numbers anymore.
                }
            }
        }

        if (this.rules.AutoGenerateMultipleRestMeasuresFromRestMeasures && this.rules.RenderMultipleRestMeasures) {
            //track number of multirests
            let beginMultiRestMeasure: SourceMeasure = undefined;
            let multiRestCount: number = 0;
            //go through all source measures again. Need to calc auto-multi-rests
            for (let idx: number = 0, len: number = musicSheet.SourceMeasures.length; idx < len; ++idx) {
                const sourceMeasure: SourceMeasure = musicSheet.SourceMeasures[idx];
                // console.log(sourceMeasure.MeasureNumber + " can be reduced: " + sourceMeasure.canBeReducedToMultiRest());
                if (!sourceMeasure.isReducedToMultiRest && sourceMeasure.canBeReducedToMultiRest()) {
                    //we've already been initialized, we are in the midst of a multirest sequence
                    if (multiRestCount > 0) {
                        beginMultiRestMeasure.isReducedToMultiRest = true;
                        beginMultiRestMeasure.multipleRestMeasureNumber = 1;
                        multiRestCount++;
                        sourceMeasure.multipleRestMeasureNumber = multiRestCount;
                        sourceMeasure.isReducedToMultiRest = true;
                        //clear out these measures. We know now that we are in multirest mode
                        for (let idx2: number = 0; idx2 < measureList[idx].length; idx2++) {
                            measureList[idx][idx2] = undefined;
                        }
                    } else { //else this is the (potential) beginning
                        beginMultiRestMeasure = sourceMeasure;
                        multiRestCount = 1;
                    }
                } else { //not multirest measure
                    if (multiRestCount > 1) { //Actual multirest sequence just happened. Process
                        beginMultiRestMeasure.multipleRestMeasures = multiRestCount;
                        //regen graphical measures for this source measure
                        const graphicalMeasures: GraphicalMeasure[] = this.createGraphicalMeasuresForSourceMeasure(
                            beginMultiRestMeasure,
                            accidentalCalculators,
                            lyricWords,
                            openOctaveShifts,
                            activeClefs
                        );
                        measureList[beginMultiRestMeasure.measureListIndex] = graphicalMeasures;
                        multiRestCount = 0;
                        beginMultiRestMeasure = undefined;
                    } else { //had a potential multirest sequence, but didn't pan out. only one measure was rests
                        multiRestCount = 0;
                        beginMultiRestMeasure = undefined;
                    }
                }
            }
            //If we reached the end of the sheet and have pending multirest measure, process
            if (multiRestCount > 1) {
                beginMultiRestMeasure.multipleRestMeasures = multiRestCount;
                beginMultiRestMeasure.isReducedToMultiRest = true;
                //regen graphical measures for this source measure
                const graphicalMeasures: GraphicalMeasure[] = this.createGraphicalMeasuresForSourceMeasure(
                    beginMultiRestMeasure,
                    accidentalCalculators,
                    lyricWords,
                    openOctaveShifts,
                    activeClefs
                );
                measureList[beginMultiRestMeasure.measureListIndex] = graphicalMeasures;
                multiRestCount = 0;
                beginMultiRestMeasure = undefined;
            }
        }

        const staffIsPercussionArray: Array<boolean> =
                        activeClefs.map(clef => (clef.ClefType === ClefEnum.percussion));

        this.handleStaffEntries(staffIsPercussionArray);
        this.calculateVerticalContainersList();
        this.setIndicesToVerticalGraphicalContainers();
    }

    /**
     * The main method for the Calculator.
     */
    public calculate(): void {
        this.musicSystems = [];

        this.clearSystemsAndMeasures();

        // delete graphicalObjects (currently: ties) that will be recalculated, newly create GraphicalObjects streching over a single StaffEntry
        this.clearRecreatedObjects();

        // this.graphicalMusicSheet.initializeActiveClefs(); // could have been changed since last render?

        this.createGraphicalTies();

        // calculate SheetLabelBoundingBoxes
        this.calculateSheetLabelBoundingBoxes();
        this.calculateXLayout(this.graphicalMusicSheet, this.maxInstrNameLabelLength());

        // create List<MusicPage>
        this.graphicalMusicSheet.MusicPages.length = 0;

        // create new MusicSystems and StaffLines (as many as necessary) and populate them with Measures from measureList
        this.calculateMusicSystems();

        // Add some white space at the end of the piece:
        //this.graphicalMusicSheet.MusicPages[0].PositionAndShape.BorderMarginBottom += 9;

        // transform Relative to Absolute Positions
        //This is called for each measure in calculate music systems (calculateLines -> calculateSkyBottomLines)
        GraphicalMusicSheet.transformRelativeToAbsolutePosition(this.graphicalMusicSheet);
    }

    public calculateXLayout(graphicalMusicSheet: GraphicalMusicSheet, maxInstrNameLabelLength: number): void {
        // for each inner List in big Measure List calculate new Positions for the StaffEntries
        // and adjust Measures sizes
        // calculate max measure length for maximum zoom in.

        // let minLength: number = 0; // currently unused
        // const maxInstructionsLength: number = this.rules.MaxInstructionsConstValue;
        if (this.graphicalMusicSheet.MeasureList.length > 0) {
            /** list of vertically ordered measures belonging to one bar */
            // let measures: GraphicalMeasure[] = this.graphicalMusicSheet.MeasureList[0];
            // let minimumStaffEntriesWidth: number = this.calculateMeasureXLayout(measures);
            // minimumStaffEntriesWidth = this.calculateMeasureWidthFromStaffEntries(measures, minimumStaffEntriesWidth);
            // MusicSheetCalculator.setMeasuresMinStaffEntriesWidth(measures, minimumStaffEntriesWidth);
            // minLength = minimumStaffEntriesWidth * 1.2 + maxInstrNameLabelLength + maxInstructionsLength;
            let maxWidth: number = 0;
            let measures: GraphicalMeasure[];
            let measureWidthFactor: number = 1;
            for (let i: number = 0; i < this.graphicalMusicSheet.MeasureList.length; i++) {
                measures = this.graphicalMusicSheet.MeasureList[i];
                let minimumStaffEntriesWidth: number = this.calculateMeasureXLayout(measures);
                minimumStaffEntriesWidth = this.calculateMeasureWidthFromStaffEntries(measures, minimumStaffEntriesWidth);
                if (minimumStaffEntriesWidth > maxWidth) {
                    maxWidth = minimumStaffEntriesWidth;
                }
                const globalWidthFactor: number = this.graphicalMusicSheet.ParentMusicSheet.MeasureWidthFactor;
                for (const verticalMeasure of measures) {
                    if (verticalMeasure?.parentSourceMeasure.WidthFactor) { // some of these GraphicalMeasures might be undefined (multi-rest)
                        measureWidthFactor = verticalMeasure.parentSourceMeasure.WidthFactor;
                        break;
                    }
                }
                minimumStaffEntriesWidth *= globalWidthFactor * measureWidthFactor;
                //console.log(`min width for measure ${measures[0].MeasureNumber}: ${minimumStaffEntriesWidth}`);
                MusicSheetCalculator.setMeasuresMinStaffEntriesWidth(measures, minimumStaffEntriesWidth);
                // minLength = Math.max(minLength, minimumStaffEntriesWidth * 1.2 + maxInstructionsLength);
            }
            if (this.rules.FixedMeasureWidth) {
                // experimental: use the same measure width for all measures
                //   here we take the maximum measure width for now,
                //   otherwise Vexflow's layout can get completely messed up and place everything on top of each other,
                //   if it gets less width than it says it needs as a minimum for a measure. (formatter.preCalculateMinTotalWidth)
                let targetWidth: number = maxWidth;
                if (this.rules.FixedMeasureWidthFixedValue) {
                    targetWidth = this.rules.FixedMeasureWidthFixedValue;
                }
                for (let i: number = 0; i < this.graphicalMusicSheet.MeasureList.length; i++) {
                    measures = this.graphicalMusicSheet.MeasureList[i];
                    if (!this.rules.FixedMeasureWidthUseForPickupMeasures && measures[0]?.parentSourceMeasure.ImplicitMeasure) {
                        // note that measures[0] is undefined for multi-measure rests
                        continue;
                    }
                    MusicSheetCalculator.setMeasuresMinStaffEntriesWidth(measures, targetWidth);
                }
            }
        }
        // this.graphicalMusicSheet.MinAllowedSystemWidth = minLength; // currently unused
    }

    public calculateMeasureWidthFromStaffEntries(measuresVertical: GraphicalMeasure[], oldMinimumStaffEntriesWidth: number): number {
        throw new Error(this.abstractNotImplementedErrorMessage);
    }

    protected formatMeasures(): void {
        throw new Error(this.abstractNotImplementedErrorMessage);
    }

    /**
     * Calculates the x layout of the staff entries within the staff measures belonging to one source measure.
     * All staff entries are x-aligned throughout all the measures.
     * @param measures - The minimum required x width of the source measure
     */
    protected calculateMeasureXLayout(measures: GraphicalMeasure[]): number {
        throw new Error(this.abstractNotImplementedErrorMessage);
    }

    /**
     * Called for every source measure when generating the list of staff measures for it.
     */
    protected initGraphicalMeasuresCreation(): void {
        throw new Error(this.abstractNotImplementedErrorMessage);
    }

    protected handleBeam(graphicalNote: GraphicalNote, beam: Beam, openBeams: Beam[]): void {
        throw new Error(this.abstractNotImplementedErrorMessage);
    }

    /**
     * Check if the tied graphical note belongs to any beams or tuplets and react accordingly.
     * @param tiedGraphicalNote
     * @param beams
     * @param activeClef
     * @param octaveShiftValue
     * @param graphicalStaffEntry
     * @param duration
     * @param openTie
     * @param isLastTieNote
     */
    protected handleTiedGraphicalNote(tiedGraphicalNote: GraphicalNote, beams: Beam[], activeClef: ClefInstruction,
                                      octaveShiftValue: OctaveEnum, graphicalStaffEntry: GraphicalStaffEntry, duration: Fraction,
                                      openTie: Tie, isLastTieNote: boolean): void {
        throw new Error(this.abstractNotImplementedErrorMessage);
    }

    protected handleVoiceEntryLyrics(voiceEntry: VoiceEntry, graphicalStaffEntry: GraphicalStaffEntry,
                                     openLyricWords: LyricWord[]): void {
        throw new Error(this.abstractNotImplementedErrorMessage);
    }

    protected handleVoiceEntryOrnaments(ornamentContainer: OrnamentContainer, voiceEntry: VoiceEntry,
                                        graphicalStaffEntry: GraphicalStaffEntry): void {
        throw new Error(this.abstractNotImplementedErrorMessage);
    }

    protected handleVoiceEntryArticulations(articulations: Articulation[],
                                            voiceEntry: VoiceEntry,
                                            staffEntry: GraphicalStaffEntry): void {
        throw new Error(this.abstractNotImplementedErrorMessage);
    }

    /**
     * Adds a technical instruction at the given staff entry.
     * @param technicalInstructions
     * @param voiceEntry
     * @param staffEntry
     */
    protected handleVoiceEntryTechnicalInstructions(technicalInstructions: TechnicalInstruction[],
                                                    voiceEntry: VoiceEntry, staffEntry: GraphicalStaffEntry): void {
        throw new Error(this.abstractNotImplementedErrorMessage);
    }


    protected handleTuplet(graphicalNote: GraphicalNote, tuplet: Tuplet, openTuplets: Tuplet[]): void {
        throw new Error(this.abstractNotImplementedErrorMessage);
    }

    protected layoutVoiceEntry(voiceEntry: VoiceEntry, graphicalNotes: GraphicalNote[],
                               graphicalStaffEntry: GraphicalStaffEntry, hasPitchedNote: boolean): void {
        throw new Error(this.abstractNotImplementedErrorMessage);
    }

    protected layoutStaffEntry(graphicalStaffEntry: GraphicalStaffEntry): void {
        throw new Error(this.abstractNotImplementedErrorMessage);
    }

    protected createGraphicalTie(tie: Tie, startGse: GraphicalStaffEntry, endGse: GraphicalStaffEntry, startNote: GraphicalNote,
                                 endNote: GraphicalNote): GraphicalTie {
        throw new Error(this.abstractNotImplementedErrorMessage);
    }

    protected updateStaffLineBorders(staffLine: StaffLine): void {
        throw new Error(this.abstractNotImplementedErrorMessage);
    }

    /**
     * Iterate through all Measures and calculates the MeasureNumberLabels.
     * @param musicSystem
     */
    protected calculateMeasureNumberPlacement(musicSystem: MusicSystem): void {
        const staffLine: StaffLine = musicSystem.StaffLines[0];
        if (!staffLine || !staffLine.Measures[0]) {
            log.warn("calculateMeasureNumberPlacement: measure undefined for system.Id " + musicSystem.Id);
            return; // TODO apparently happens in script sometimes (mp #70)
        }
        let previousMeasureNumber: number = staffLine.Measures[0].MeasureNumber;
        let labelOffsetX: number = 0;
        for (let i: number = 0; i < staffLine.Measures.length; i++) {
            const measure: GraphicalMeasure = staffLine.Measures[i];
            let skip: boolean = this.rules.RenderMeasureNumbersOnlyAtSystemStart && i > 1;
            if (i === 1 && staffLine.Measures[0].parentSourceMeasure.ImplicitMeasure) {
                skip = false; // if the first measure (i=0) is a pickup measure, we shouldn't skip measure number 1 (i=1)
            }
            if (skip) {
                return; // no more measures number labels need to be rendered for this system, so we can just return instead of continue.
            }
            if (measure.MeasureNumber === 0 || measure.MeasureNumber === 1) {
                previousMeasureNumber = measure.MeasureNumber;
                // for the first measure, this label still needs to be created. Afterwards, this variable will hold the previous label's measure number.
            }
            if (measure !== staffLine.Measures[0] && this.rules.MeasureNumberLabelXOffset) {
                labelOffsetX = this.rules.MeasureNumberLabelXOffset;
            } else {
                labelOffsetX = 0; // don't offset label for first measure in staffline
            }

            const isFirstMeasureAndNotPrintedOne: boolean = this.rules.UseXMLMeasureNumbers &&
                measure.MeasureNumber === 1 && measure.parentSourceMeasure.getPrintedMeasureNumber() !== 1;
            if ((measure.MeasureNumber === previousMeasureNumber ||
                measure.MeasureNumber >= previousMeasureNumber + this.rules.MeasureNumberLabelOffset) &&
                !measure.parentSourceMeasure.ImplicitMeasure ||
                isFirstMeasureAndNotPrintedOne) {
                if (measure.MeasureNumber !== 1 ||
                    (measure.MeasureNumber === 1 && measure !== staffLine.Measures[0]) ||
                    isFirstMeasureAndNotPrintedOne
                    ) {
                    this.calculateSingleMeasureNumberPlacement(measure, staffLine, musicSystem, labelOffsetX);
                }
                previousMeasureNumber = measure.MeasureNumber;
            }
        }
    }

    /// <summary>
    /// This method calculates a single MeasureNumberLabel and adds it to the graphical label list of the music system
    /// </summary>
    /// <param name="measure"></param>
    /// <param name="staffLine"></param>
    /// <param name="musicSystem"></param>
    private calculateSingleMeasureNumberPlacement(measure: GraphicalMeasure, staffLine: StaffLine, musicSystem: MusicSystem,
                                                  labelOffsetX: number = 0): void {
        const labelNumber: string = measure.parentSourceMeasure.getPrintedMeasureNumber().toString();
        const label: Label = new Label(labelNumber);
        label.fontFamily = this.rules.MeasureNumberFontFamily;
        // maybe give rules as argument instead of just setting fontStyle and maybe other settings manually afterwards
        const graphicalLabel: GraphicalLabel = new GraphicalLabel(label, this.rules.MeasureNumberLabelHeight,
                                                                  TextAlignmentEnum.LeftBottom, this.rules);
        graphicalLabel.ParentMeasure = measure;

        const skyBottomLineCalculator: SkyBottomLineCalculator = staffLine.SkyBottomLineCalculator;

        // calculate LabelBoundingBox and set PSI parent
        graphicalLabel.setLabelPositionAndShapeBorders();
        graphicalLabel.PositionAndShape.Parent = musicSystem.PositionAndShape;

        // calculate relative Position
        const relativeX: number = staffLine.PositionAndShape.RelativePosition.x +
            measure.PositionAndShape.RelativePosition.x - graphicalLabel.PositionAndShape.BorderMarginLeft +
            labelOffsetX;
        let relativeY: number;

        // and the corresponding SkyLine indices
        let start: number = relativeX;
        let end: number = relativeX - graphicalLabel.PositionAndShape.BorderLeft + graphicalLabel.PositionAndShape.BorderRight;

        start -= staffLine.PositionAndShape.RelativePosition.x;
        end -= staffLine.PositionAndShape.RelativePosition.x;

        // correct for hypersensitive collision checks, notes having skyline extend too far to left and right
        const startCollisionCheck: number = start + 0.5;
        const endCollisionCheck: number = end - 0.5;

        // get the minimum corresponding SkyLine value
        const skyLineMinValue: number = skyBottomLineCalculator.getSkyLineMinInRange(startCollisionCheck, endCollisionCheck);

        if (measure === staffLine.Measures[0]) {
            // must take into account possible MusicSystem Brackets
            let minBracketTopBorder: number = 0;
            if (musicSystem.GroupBrackets.length > 0) {
                for (const groupBracket of musicSystem.GroupBrackets) {
                    minBracketTopBorder = Math.min(minBracketTopBorder, groupBracket.PositionAndShape.BorderTop);
                }
            } else if (measure.ParentStaff.ParentInstrument.Parent) { // Parent InstrumentalGroup
                // note that GroupBracket creation is currently done after measure number creation, so we have to check it indirectly.
                minBracketTopBorder = -1;
            }
            relativeY = Math.min(skyLineMinValue, minBracketTopBorder);
        } else {
            relativeY = skyLineMinValue;
        }

        const maxMeasureNumberDistanceAboveStaff: number = this.rules.MeasureNumberLabelHeight + 0.7;
        relativeY = Math.max(relativeY, -maxMeasureNumberDistanceAboveStaff);
        relativeY = Math.min(0, relativeY);

        graphicalLabel.PositionAndShape.RelativePosition = new PointF2D(relativeX, relativeY);
        musicSystem.MeasureNumberLabels.push(graphicalLabel);
    }
    //So we can apply slurs first, then do these
    private calculateMeasureNumberSkyline(musicSystem: MusicSystem): void {
        const staffLine: StaffLine = musicSystem.StaffLines[0];
        for(const measureNumberLabel of musicSystem.MeasureNumberLabels) {
            // and the corresponding SkyLine indices
            let start: number = measureNumberLabel.PositionAndShape.RelativePosition.x;
            let end: number = start - measureNumberLabel.PositionAndShape.BorderLeft + measureNumberLabel.PositionAndShape.BorderRight;
            start -= staffLine.PositionAndShape.RelativePosition.x;
            end -= staffLine.PositionAndShape.RelativePosition.x;
            staffLine.SkyBottomLineCalculator.updateSkyLineInRange(start, end,
                measureNumberLabel.PositionAndShape.RelativePosition.y + measureNumberLabel.PositionAndShape.BorderMarginTop);
        }
    }

    /**
     * Calculate the shape (Bézier curve) for this tie.
     * @param tie
     * @param tieIsAtSystemBreak
     */
    protected layoutGraphicalTie(tie: GraphicalTie, tieIsAtSystemBreak: boolean, isTab: boolean): void {
        throw new Error(this.abstractNotImplementedErrorMessage);
    }

    /**
     * Calculate the Lyrics YPositions for a single [[StaffLine]].
     * @param staffLine
     * @param lyricVersesNumber
     */
    protected calculateSingleStaffLineLyricsPosition(staffLine: StaffLine, lyricVersesNumber: string[]): GraphicalStaffEntry[] {
        let numberOfVerses: number = 0;
        let lyricsStartYPosition: number = this.rules.StaffHeight; // Add offset to prevent collision
        const relevantVerseNumbers: Map<string, boolean> = new Map<string, boolean>();
        const lyricsStaffEntriesList: GraphicalStaffEntry[] = [];
        const skyBottomLineCalculator: SkyBottomLineCalculator = staffLine.SkyBottomLineCalculator;

        // first find maximum Ycoordinate for the whole StaffLine
        let len: number = staffLine.Measures.length;
        for (let idx: number = 0; idx < len; ++idx) {
            const measure: GraphicalMeasure = staffLine.Measures[idx];
            const measureRelativePosition: PointF2D = measure.PositionAndShape.RelativePosition;
            const len2: number = measure.staffEntries.length;
            for (let idx2: number = 0; idx2 < len2; ++idx2) {
                const staffEntry: GraphicalStaffEntry = measure.staffEntries[idx2];

                // Collect relevant verse numbers
                const len3: number = staffEntry.LyricsEntries.length;
                for (let idx3: number = 0; idx3 < len3; ++idx3) {
                    const lyricsEntry: LyricsEntry = staffEntry.LyricsEntries[idx3].LyricsEntry;
                    relevantVerseNumbers[lyricsEntry.VerseNumber] = lyricsEntry.IsChorus;
                }

                if (len3 > 0) {
                    lyricsStaffEntriesList.push(staffEntry);
                    numberOfVerses = Math.max(numberOfVerses, staffEntry.LyricsEntries.length);

                    // Position of Staffentry relative to StaffLine
                    const staffEntryPositionX: number = staffEntry.PositionAndShape.RelativePosition.x +
                        measureRelativePosition.x;

                    let minMarginLeft: number = Number.MAX_VALUE;
                    let maxMarginRight: number = Number.MIN_VALUE;

                    // if more than one LyricEntry in StaffEntry, find minMarginLeft, maxMarginRight of all corresponding Labels
                    for (let i: number = 0; i < staffEntry.LyricsEntries.length; i++) {
                        const lyricsEntryLabel: GraphicalLabel = staffEntry.LyricsEntries[i].GraphicalLabel;
                        minMarginLeft = Math.min(minMarginLeft, staffEntryPositionX + lyricsEntryLabel.PositionAndShape.BorderMarginLeft);
                        maxMarginRight = Math.max(maxMarginRight, staffEntryPositionX + lyricsEntryLabel.PositionAndShape.BorderMarginRight);
                    }

                    // check BottomLine in this range and take the maximum between the two values
                    const bottomLineMax: number = skyBottomLineCalculator.getBottomLineMaxInRange(minMarginLeft, maxMarginRight);
                    lyricsStartYPosition = Math.max(lyricsStartYPosition, bottomLineMax + this.rules.LyricsYMarginToBottomLine);
                }
            }
        }

        let maxPosition: number = 0;
        // iterate again through the Staffentries with LyricEntries
        len = lyricsStaffEntriesList.length;
        for (const staffEntry of lyricsStaffEntriesList) {

            // Filter verse numbers
            const filteredLyricVersesNumber: string[] = [];
            let isChorus: boolean = true;
            for (let i: number = 0; i < staffEntry.LyricsEntries.length; i++) {
                isChorus &&= staffEntry.LyricsEntries[i].LyricsEntry.IsChorus;
            }
            for (const lyricVerseNumber of lyricVersesNumber){
                if (relevantVerseNumbers[lyricVerseNumber] === isChorus) {
                    filteredLyricVersesNumber.push(lyricVerseNumber);
                }
            }

            // set LyricEntryLabel RelativePosition
            for (let i: number = 0; i < staffEntry.LyricsEntries.length; i++) {
                const lyricEntry: GraphicalLyricEntry = staffEntry.LyricsEntries[i];
                const lyricsEntryLabel: GraphicalLabel = lyricEntry.GraphicalLabel;

                // read the verseNumber and get index of this number in the sorted LyricVerseNumbersList of Instrument
                // eg verseNumbers: 2,3,4,6 => 1,2,3,4
                const verseNumber: string = lyricEntry.LyricsEntry.VerseNumber;
                const sortedLyricVerseNumberIndex: number = filteredLyricVersesNumber.indexOf(verseNumber);
                const firstPosition: number = lyricsStartYPosition + this.rules.LyricsHeight + this.rules.VerticalBetweenLyricsDistance +
                    this.rules.LyricsYOffsetToStaffHeight;

                // Y-position calculated according to aforementioned mapping
                const position: number = firstPosition + (this.rules.VerticalBetweenLyricsDistance + this.rules.LyricsHeight) * sortedLyricVerseNumberIndex;
                // TODO not sure what this leadsheet lyrics positioning was supposed to be, but it seems to ALWAYS put the lyrics inside the stafflines now.
                // if (this.leadSheet) {
                //     position = 3.4 + (this.rules.VerticalBetweenLyricsDistance + this.rules.LyricsHeight) * (sortedLyricVerseNumberIndex);
                // }
                const previousRelativeX: number = lyricsEntryLabel.PositionAndShape.RelativePosition.x;
                lyricsEntryLabel.PositionAndShape.RelativePosition = new PointF2D(previousRelativeX, position);
                lyricsEntryLabel.Label.fontStyle = lyricEntry.LyricsEntry.FontStyle;
                maxPosition = Math.max(maxPosition, position);
            }
        }

        // update BottomLine (on the whole StaffLine's length)
        if (lyricsStaffEntriesList.length > 0) {
            const endX: number = staffLine.PositionAndShape.Size.width;
            let startX: number = lyricsStaffEntriesList[0].PositionAndShape.RelativePosition.x +
                lyricsStaffEntriesList[0].PositionAndShape.BorderMarginLeft +
                lyricsStaffEntriesList[0].parentMeasure.PositionAndShape.RelativePosition.x;
            startX = startX > endX ? endX : startX;
            skyBottomLineCalculator.updateBottomLineInRange(startX, endX, maxPosition);
        }
        return lyricsStaffEntriesList;
    }

    /**
     * calculates the dashes of lyric words and the extending underscore lines of syllables sung on more than one note.
     * @param lyricsStaffEntries
     */
    protected calculateLyricsExtendsAndDashes(lyricsStaffEntries: GraphicalStaffEntry[]): void {
        // iterate again to create now the extend lines and dashes for words
        for (let idx: number = 0, len: number = lyricsStaffEntries.length; idx < len; ++idx) {
            const staffEntry: GraphicalStaffEntry = lyricsStaffEntries[idx];
            // set LyricEntryLabel RelativePosition
            for (let i: number = 0; i < staffEntry.LyricsEntries.length; i++) {
                const lyricEntry: GraphicalLyricEntry = staffEntry.LyricsEntries[i];
                // calculate LyricWord's Dashes and underscoreLine
                if (lyricEntry.ParentLyricWord &&
                    lyricEntry.ParentLyricWord.GraphicalLyricsEntries[lyricEntry.ParentLyricWord.GraphicalLyricsEntries.length - 1] !== lyricEntry) {
                    this.calculateSingleLyricWord(lyricEntry);
                }
                // calculate the underscore line extend if needed
                if (lyricEntry.LyricsEntry.extend) {
                    this.calculateLyricExtend(lyricEntry);
                }
            }
        }
    }

    /**
     * Calculate a single OctaveShift for a [[MultiExpression]].
     * @param sourceMeasure
     * @param multiExpression
     * @param measureIndex
     * @param staffIndex
     */
    protected calculateSingleOctaveShift(sourceMeasure: SourceMeasure, multiExpression: MultiExpression,
                                         measureIndex: number, staffIndex: number): void {
        throw new Error(this.abstractNotImplementedErrorMessage);
    }

    /**
     * Calculate a single Pedal for a [[MultiExpression]].
     * @param sourceMeasure
     * @param multiExpression
     * @param measureIndex
     * @param staffIndex
     */
    protected abstract calculateSinglePedal(sourceMeasure: SourceMeasure, multiExpression: MultiExpression,
        measureIndex: number, staffIndex: number): void;

    /**
     * Calculate a single Wavy Line for a [[MultiExpression]].
     * @param sourceMeasure
     * @param multiExpression
     * @param measureIndex
     * @param staffIndex
     */
    protected abstract calculateSingleWavyLine(sourceMeasure: SourceMeasure, multiExpression: MultiExpression,
        measureIndex: number, staffIndex: number): void;

    /**
     * Calculate all the textual [[RepetitionInstruction]]s (e.g. dal segno) for a single [[SourceMeasure]].
     * @param repetitionInstruction
     * @param measureIndex
     */
    protected calculateWordRepetitionInstruction(repetitionInstruction: RepetitionInstruction,
                                                 measureIndex: number): void {
        throw new Error(this.abstractNotImplementedErrorMessage);
    }

    /**
     * Calculate all the Mood and Unknown Expressions for a single [[MultiExpression]].
     * @param multiExpression
     * @param measureIndex
     * @param staffIndex
     */
    protected calculateMoodAndUnknownExpression(multiExpression: MultiExpression, measureIndex: number, staffIndex: number): void {
        // calculate absolute Timestamp
        const absoluteTimestamp: Fraction = multiExpression.AbsoluteTimestamp;
        const measures: GraphicalMeasure[] = this.graphicalMusicSheet.MeasureList[measureIndex];
        let relative: PointF2D = new PointF2D();

        const defaultYXml: number = multiExpression.UnknownList[0]?.defaultYXml;
        if ((multiExpression.MoodList.length > 0) || (multiExpression.UnknownList.length > 0)) {
        let combinedExprString: string  = "";
        for (let idx: number = 0, len: number = multiExpression.EntriesList.length; idx < len; ++idx) {
            const entry: MultiExpressionEntry = multiExpression.EntriesList[idx];
            if (entry.prefix !== "") {
                if (combinedExprString === "") {
                    combinedExprString += entry.prefix;
                } else {
                    combinedExprString += " " + entry.prefix;
                }
            }
            if (combinedExprString === "") {
                combinedExprString += entry.label;
            } else {
                combinedExprString += " " + entry.label;
            }
        }
        const staffLine: StaffLine = measures[staffIndex].ParentStaffLine;
        if (!staffLine) {
            log.debug("MusicSheetCalculator.calculateMoodAndUnknownExpression: staffLine undefined. Returning.");
            return;
        }
        relative = this.getRelativePositionInStaffLineFromTimestamp(absoluteTimestamp, staffIndex, staffLine, staffLine?.isPartOfMultiStaffInstrument());

        if (Math.abs(relative.x - 0) < 0.0001) {
            relative.x = measures[staffIndex].beginInstructionsWidth + this.rules.RhythmRightMargin;
        }

        const fontHeight: number = this.rules.UnknownTextHeight;
        const placement: PlacementEnum = multiExpression.getPlacementOfFirstEntry();
        const graphLabel: GraphicalLabel  = this.calculateLabel(staffLine,
                                                                relative, combinedExprString,
                                                                multiExpression.getFontstyleOfFirstEntry(),
                                                                placement,
                                                                fontHeight);
        const colorXML: string = multiExpression.getColorXMLOfFirstEntry();
        if (this.rules.ExpressionsUseXMLColor && colorXML) {
            graphLabel.ColorXML = colorXML;
        }
        if (this.rules.PlaceWordsInsideStafflineFromXml) {
            if (defaultYXml < 0 && defaultYXml > -50) { // within staffline
                let newY: number = defaultYXml / 10; // OSMD units
                newY += this.rules.PlaceWordsInsideStafflineYOffset;
                graphLabel.PositionAndShape.RelativePosition.y = newY;
            }
        }

        const gue: GraphicalUnknownExpression = new GraphicalUnknownExpression(
            staffLine, graphLabel, placement, measures[staffIndex]?.parentSourceMeasure, multiExpression);
        //    multiExpression); // TODO would be nice to hand over and save reference to original expression,
        //                         but MultiExpression is not an AbstractExpression.
        staffLine.AbstractExpressions.push(gue);
        }
    }

    /**
     * Delete all Objects that must be recalculated.
     * If graphicalMusicSheet.reCalculate has been called, then this method will be called to reset or remove all flexible
     * graphical music symbols (e.g. Ornaments, Lyrics, Slurs) graphicalMusicSheet will have MusicPages, they will have MusicSystems etc...
     */
    protected clearRecreatedObjects(): void {
        // Clear StaffEntries with GraphicalTies
        for (let idx: number = 0, len: number = this.staffEntriesWithGraphicalTies.length; idx < len; ++idx) {
            const staffEntriesWithGraphicalTie: GraphicalStaffEntry = this.staffEntriesWithGraphicalTies[idx];
            staffEntriesWithGraphicalTie.GraphicalTies.length = 0;
        }
        this.staffEntriesWithGraphicalTies.length = 0;
        return;
    }

    /**
     * This method handles a [[StaffEntryLink]].
     * @param graphicalStaffEntry
     * @param staffEntryLinks
     */
    protected handleStaffEntryLink(graphicalStaffEntry: GraphicalStaffEntry,
                                   staffEntryLinks: StaffEntryLink[]): void {
        log.debug("handleStaffEntryLink not implemented");
    }

    /**
     * Store the newly computed [[Measure]]s in newly created [[MusicSystem]]s.
     */
    protected calculateMusicSystems(): void {
        if (!this.graphicalMusicSheet.MeasureList) {
            return;
        }

        const allMeasures: GraphicalMeasure[][] = this.graphicalMusicSheet.MeasureList;
        if (!allMeasures) {
            return;
        }
        if (this.rules.MinMeasureToDrawIndex > allMeasures.length - 1) {
            log.debug("minimum measure to draw index out of range. resetting min measure index to limit.");
            this.rules.MinMeasureToDrawIndex = allMeasures.length - 1;
        }

        // visible 2D-MeasureList
        const visibleMeasureList: GraphicalMeasure[][] = [];
        for (let idx: number = this.rules.MinMeasureToDrawIndex, len: number = allMeasures.length;
            idx < len && idx <= this.rules.MaxMeasureToDrawIndex; ++idx) {
            const graphicalMeasures: GraphicalMeasure[] = allMeasures[idx];
            const visiblegraphicalMeasures: GraphicalMeasure[] = [];
            for (let idx2: number = 0, len2: number = graphicalMeasures.length; idx2 < len2; ++idx2) {
                const graphicalMeasure: GraphicalMeasure = allMeasures[idx][idx2];

                if (graphicalMeasure?.isVisible()) {
                    visiblegraphicalMeasures.push(graphicalMeasure);

                    if (this.rules.ColoringEnabled) {
                        // (re-)color notes
                        for (const staffEntry of graphicalMeasure.staffEntries) {
                            for (const gve of staffEntry.graphicalVoiceEntries) {
                                gve.applyCustomNoteheads();
                                gve.color();
                            }
                        }
                    }
                }
            }
            visibleMeasureList.push(visiblegraphicalMeasures);
        }

        // find out how many StaffLine Instances we need
        let numberOfStaffLines: number = 0;

        for (let idx: number = 0, len: number = visibleMeasureList.length; idx < len; ++idx) {
            const gmlist: GraphicalMeasure[] = visibleMeasureList[idx];
            numberOfStaffLines = Math.max(gmlist.length, numberOfStaffLines);

            break;
        }
        if (numberOfStaffLines === 0) {
            return;
        }


        // build the MusicSystems (and StaffLines)
        const musicSystemBuilder: MusicSystemBuilder = new MusicSystemBuilder();
        musicSystemBuilder.initialize(this.graphicalMusicSheet, visibleMeasureList, numberOfStaffLines);
        this.musicSystems = musicSystemBuilder.buildMusicSystems();

        this.formatMeasures();

        // check for Measures with only WholeRestNotes and correct their X-Position (middle of Measure)
        // this.checkMeasuresForWholeRestNotes(); // this currently does nothing
        if (!this.leadSheet) {
            // calculate Beam Placement
            // this.calculateBeams(); // does nothing for now, because layoutBeams() is an empty method
            // possible Displacement of RestNotes
            this.optimizeRestPlacement();
            // possible Displacement of RestNotes
            this.calculateStaffEntryArticulationMarks();
            if (this.rules.RenderSlurs) { // technically we should separate slurs and ties, but shouldn't be relevant for now
                // calculate Ties
                this.calculateTieCurves();
            }
        }
        // calculate Sky- and BottomLine
        // will have reasonable values only between ObjectsBorders (eg StaffEntries)
        this.calculateSkyBottomLines();
        // calculate TupletsNumbers
        this.calculateTupletNumbers();

        // calculate MeasureNumbers
        if (this.rules.RenderMeasureNumbers) {
            for (let idx: number = 0, len: number = this.musicSystems.length; idx < len; ++idx) {
                const musicSystem: MusicSystem = this.musicSystems[idx];
                this.calculateMeasureNumberPlacement(musicSystem);
            }
        }
        if (this.rules.RenderFingerings) {
            this.calculateFingerings(); // if this is done after slurs, fingerings can be on top of slurs
        }
        // calculate Slurs
        if (!this.leadSheet && this.rules.RenderSlurs) {
            this.calculateSlurs();
        }
        this.calculateGlissandi();
        //Calculate measure number skyline AFTER slurs
        if (this.rules.RenderMeasureNumbers) {
            for (let idx: number = 0, len: number = this.musicSystems.length; idx < len; ++idx) {
                const musicSystem: MusicSystem = this.musicSystems[idx];
                this.calculateMeasureNumberSkyline(musicSystem);
            }
        }
        // calculate StaffEntry Ornaments
        // (must come after Slurs)
        if (!this.leadSheet) {
            this.calculateOrnaments();
        }
        // calculate StaffEntry ChordSymbols
        this.calculateChordSymbols();
        if (!this.leadSheet) {
            // calculate all Instantaneous/Continuous Dynamics Expressions
            this.calculateDynamicExpressions();
            // calculate all Mood and Unknown Expression
            this.calculateMoodAndUnknownExpressions();
            // Calculate the alignment of close expressions
            this.calculateExpressionAlignements();
            // calculate all OctaveShifts
            this.calculateOctaveShifts();
            if (this.rules.RenderPedals) {
                // calculate all Pedal Expressions
                this.calculatePedals();
            }
            //calculate all wavy lines (vibrato, trill marks)
            if (this.rules.RenderWavyLines) {
                this.calculateWavyLines();
            }
            // calculate RepetitionInstructions (Dal Segno, Coda, etc)
            this.calculateWordRepetitionInstructions();
        }
        // calculate endings last, so they appear above measure numbers
        this.calculateRepetitionEndings();
        // calcualte all Tempo Expressions
        if (!this.leadSheet) {
            this.calculateTempoExpressions();
        }
        this.calculateRehearsalMarks();

        // calculate all LyricWords Positions
        this.calculateLyricsPosition();

        // update all StaffLine's Borders
        // create temporary Object, just to call the methods (in order to avoid declaring them static)
        for (let idx2: number = 0, len2: number = this.musicSystems.length; idx2 < len2; ++idx2) {
            const musicSystem: MusicSystem = this.musicSystems[idx2];
            for (let idx3: number = 0, len3: number = musicSystem.StaffLines.length; idx3 < len3; ++idx3) {
                const staffLine: StaffLine = musicSystem.StaffLines[idx3];
                this.updateStaffLineBorders(staffLine);
            }
        }

        // calculate Y-spacing -> MusicPages are created here
        musicSystemBuilder.calculateSystemYLayout();
        // calculate Comments for each Staffline
        this.calculateComments();
        // calculate marked Areas for Systems
        this.calculateMarkedAreas();

        // the following must be done after Y-spacing, when the MusicSystems's final Dimensions are set
        // set the final yPositions of Objects such as SystemLabels and SystemLinesContainers,
        // create all System Lines, Brackets and MeasureNumbers (for all systems and for all pages)
        for (let idx: number = 0, len: number = this.graphicalMusicSheet.MusicPages.length; idx < len; ++idx) {
            const graphicalMusicPage: GraphicalMusicPage = this.graphicalMusicSheet.MusicPages[idx];
            for (let idx2: number = 0, len2: number = graphicalMusicPage.MusicSystems.length; idx2 < len2; ++idx2) {
                const isFirstSystem: boolean = idx === 0 && idx2 === 0;
                const musicSystem: MusicSystem = graphicalMusicPage.MusicSystems[idx2];
                musicSystem.setMusicSystemLabelsYPosition();
                if (!this.leadSheet) {
                    musicSystem.setYPositionsToVerticalLineObjectsAndCreateLines(this.rules);
                    musicSystem.createSystemLeftLine(this.rules.SystemThinLineWidth, this.rules.SystemLabelsRightMargin, isFirstSystem);
                    musicSystem.createInstrumentBrackets(this.graphicalMusicSheet.ParentMusicSheet.Instruments, this.rules.StaffHeight);
                    musicSystem.createGroupBrackets(this.graphicalMusicSheet.ParentMusicSheet.InstrumentalGroups, this.rules.StaffHeight, 0);
                    musicSystem.alignBeginInstructions();
                } else if (musicSystem === musicSystem.Parent.MusicSystems[0]) {
                    musicSystem.createSystemLeftLine(this.rules.SystemThinLineWidth, this.rules.SystemLabelsRightMargin, isFirstSystem);
                }
                musicSystem.calculateBorders(this.rules);
            }
            const distance: number = graphicalMusicPage.MusicSystems[0].PositionAndShape.BorderTop;
            for (let idx2: number = 0, len2: number = graphicalMusicPage.MusicSystems.length; idx2 < len2; ++idx2) {
                const musicSystem: MusicSystem = graphicalMusicPage.MusicSystems[idx2];
                // let newPosition: PointF2D = new PointF2D(musicSystem.PositionAndShape.RelativePosition.x,
                // musicSystem.PositionAndShape.RelativePosition.y - distance);
                musicSystem.PositionAndShape.RelativePosition =
                    new PointF2D(musicSystem.PositionAndShape.RelativePosition.x, musicSystem.PositionAndShape.RelativePosition.y - distance);
            }
            // add ActivitySymbolClickArea - currently unused, extends boundingbox of MusicSystem unnecessarily -> PageRightMargin 0 impossible
            // for (let idx2: number = 0, len2: number = graphicalMusicPage.MusicSystems.length; idx2 < len2; ++idx2) {
            //     const musicSystem: MusicSystem = graphicalMusicPage.MusicSystems[idx2];
            //     for (let idx3: number = 0, len3: number = musicSystem.StaffLines.length; idx3 < len3; ++idx3) {
            //         const staffLine: StaffLine = musicSystem.StaffLines[idx3];
            //         staffLine.addActivitySymbolClickArea();
            //     }
            // }

            // calculate TopBottom Borders for all elements recursively
            //   necessary for composer label (page labels) for high notes in first system
            graphicalMusicPage.PositionAndShape.calculateTopBottomBorders();
            // TODO how much performance does this cost? can we reduce the amount of calculations, e.g. only checking top?

            // calculate all Labels's Positions for the first Page
            if (graphicalMusicPage === this.graphicalMusicSheet.MusicPages[0]) {
                this.calculatePageLabels(graphicalMusicPage);
            }

            // calculate TopBottom Borders for all elements recursively
            graphicalMusicPage.PositionAndShape.calculateTopBottomBorders(); // this is where top bottom borders were originally calculated (only once)
        }
    }

    protected calculateMarkedAreas(): void {
        //log.debug("calculateMarkedAreas not implemented");
        return;
    }

    protected calculateComments(): void {
        //log.debug("calculateComments not implemented");
        return;
    }

    protected calculateChordSymbols(): void {
        for (const musicSystem of this.musicSystems) {
            for (const staffLine of musicSystem.StaffLines) {
                const skybottomcalculator: SkyBottomLineCalculator = staffLine.SkyBottomLineCalculator;
                let minimumOffset: number = Number.MAX_SAFE_INTEGER; // only calculated if option set
                let maximumOffset: number = Number.MIN_SAFE_INTEGER;
                if (this.rules.ChordSymbolYAlignment && this.rules.ChordSymbolYAlignmentScope === "staffline") {
                    // get the max y position of all chord symbols in the staffline in advance
                    const alignmentScopedStaffEntries: GraphicalStaffEntry[] = [];
                    for (const measure of staffLine.Measures) {
                        alignmentScopedStaffEntries.push(...measure.staffEntries);
                    }
                    const { minOffset, maxOffset } = this.calculateAlignedChordSymbolsOffset(alignmentScopedStaffEntries, skybottomcalculator);
                    minimumOffset = minOffset;
                    maximumOffset = maxOffset;
                }
                for (let measureStafflineIndex: number = 0; measureStafflineIndex < staffLine.Measures.length; measureStafflineIndex++) {
                    const measure: GraphicalMeasure = staffLine.Measures[measureStafflineIndex];
                    if (this.rules.ChordSymbolYAlignment && this.rules.ChordSymbolYAlignmentScope === "measure") {
                        const { minOffset, maxOffset } = this.calculateAlignedChordSymbolsOffset(measure.staffEntries, skybottomcalculator);
                        minimumOffset = minOffset;
                        maximumOffset = maxOffset;
                    }
                    let previousChordContainer: GraphicalChordSymbolContainer;
                    for (const staffEntry of measure.staffEntries) {
                        if (!staffEntry.graphicalChordContainers || staffEntry.graphicalChordContainers.length === 0) {
                            continue;
                        }
                        for (let i: number = 0; i < staffEntry.graphicalChordContainers.length; i++) {
                            const graphicalChordContainer: GraphicalChordSymbolContainer = staffEntry.graphicalChordContainers[i];
                            // check for chord not over a note
                            if (staffEntry.graphicalVoiceEntries.length === 0 && staffEntry.relInMeasureTimestamp.RealValue > 0) {
                                // re-position (second chord symbol on whole measure rest)
                                let firstNoteStartX: number = 0;
                                if (measure.staffEntries[0].relInMeasureTimestamp.RealValue === 0) {
                                    firstNoteStartX = measure.staffEntries[0].PositionAndShape.RelativePosition.x;
                                    if (measure.MeasureNumber === 1) {
                                        firstNoteStartX += this.rules.ChordSymbolWholeMeasureRestXOffsetMeasure1;
                                        // shift second chord same way as first chord
                                    }
                                }
                                const measureEndX: number = measure.PositionAndShape.Size.width - measure.endInstructionsWidth;
                                const proportionInMeasure: number = staffEntry.relInMeasureTimestamp.RealValue / measure.parentSourceMeasure.Duration.RealValue;
                                let newStartX: number = firstNoteStartX + (measureEndX - firstNoteStartX) * proportionInMeasure +
                                    graphicalChordContainer.PositionAndShape.BorderMarginLeft; // negative -> shift a bit left to where it starts visually
                                if (previousChordContainer) {
                                    // prevent overlap to previous chord symbol
                                    newStartX = Math.max(newStartX, previousChordContainer.PositionAndShape.RelativePosition.x +
                                        previousChordContainer.GraphicalLabel.PositionAndShape.Size.width +
                                        this.rules.ChordSymbolXSpacing);
                                }
                                graphicalChordContainer.PositionAndShape.RelativePosition.x = newStartX;
                                graphicalChordContainer.PositionAndShape.Parent = measure.staffEntries[0].PositionAndShape.Parent;
                                // TODO it would be more clean to set the staffEntry relative position instead of the container's,
                                //   so that the staff entry also gets a valid position (and not relative 0),
                                //   but this is tricky with elongationFactor, skyline etc, would need some adjustments
                                // // graphicalChordContainer.PositionAndShape.Parent = measure.staffEntries[0].PositionAndShape.Parent; // not here
                                // //   don't switch parent from StaffEntry if setting staffEntry.x
                                // staffEntry.PositionAndShape.RelativePosition.x = newStartX;
                                // staffEntry.PositionAndShape.calculateAbsolutePosition();
                            }
                            const gps: BoundingBox = graphicalChordContainer.PositionAndShape;
                            const parentBbox: BoundingBox = gps.Parent; // usually the staffEntry (bbox), but sometimes measure (for whole measure rests)
                            if (parentBbox.DataObject instanceof GraphicalMeasure) {
                                if (staffEntry.relInMeasureTimestamp.RealValue === 0) {
                                    gps.RelativePosition.x = Math.max(measure.beginInstructionsWidth, gps.RelativePosition.x);
                                    // beginInstructionsWidth wasn't set correctly before this
                                    if (measure.MeasureNumber === 1 && gps.RelativePosition.x > 3) {
                                        gps.RelativePosition.x += this.rules.ChordSymbolWholeMeasureRestXOffsetMeasure1;
                                    }
                                }
                            }
                            // check if there already exists a vertical staffentry with the same relative timestamp,
                            //   use its relativePosition (= x-align chord symbols to vertical staffentries in other measures)
                            if (staffEntry.PositionAndShape.RelativePosition.x === 0) {
                                const verticalMeasures: GraphicalMeasure[] = musicSystem.GraphicalMeasures[measureStafflineIndex];
                                for (const verticalMeasure of verticalMeasures) {
                                    let positionFound: boolean = false;
                                    for (const verticalSe of verticalMeasure.staffEntries) {
                                        if (verticalSe.relInMeasureTimestamp === staffEntry.relInMeasureTimestamp &&
                                            verticalSe.PositionAndShape.RelativePosition.x !== 0) {
                                            gps.RelativePosition.x = verticalSe.PositionAndShape.RelativePosition.x;
                                            positionFound = true;
                                            break;
                                        }
                                    }
                                    if (positionFound) {
                                        break;
                                    }
                                }
                            }
                            const start: number = gps.BorderMarginLeft + parentBbox.AbsolutePosition.x + gps.RelativePosition.x;
                            const end: number = gps.BorderMarginRight + parentBbox.AbsolutePosition.x + gps.RelativePosition.x;
                            const placement: PlacementEnum = graphicalChordContainer.GetChordSymbolContainer.Placement;
                            if (placement === PlacementEnum.Below) {
                                if (!this.rules.ChordSymbolYAlignment || maximumOffset < 0) {
                                    maximumOffset = skybottomcalculator.getBottomLineMaxInRange(start, end);
                                }
                            } else if (placement === PlacementEnum.Above) {
                                if (!this.rules.ChordSymbolYAlignment || minimumOffset > 0) {
                                    //minimumOffset = this.calculateAlignedChordSymbolsOffset([staffEntry], skybottomcalculator);
                                    minimumOffset = skybottomcalculator.getSkyLineMinInRange(start, end); // same as above, less code executed
                                }
                            }
                            let yShift: number = 0;
                            if (i === 0) {
                                yShift += this.rules.ChordSymbolYOffset;
                                yShift += 0.1; // above is a bit closer to the notes than below ones for some reason
                            } else {
                                yShift += this.rules.ChordSymbolYPadding;
                            }
                            if (placement !== PlacementEnum.Below) {
                                yShift *= -1;
                            }
                            const gLabel: GraphicalLabel = graphicalChordContainer.GraphicalLabel;
                            if (placement === PlacementEnum.Below) {
                                gLabel.PositionAndShape.RelativePosition.y = maximumOffset + yShift;
                                gLabel.setLabelPositionAndShapeBorders();
                                gLabel.PositionAndShape.calculateBoundingBox();
                                skybottomcalculator.updateBottomLineInRange(start, end,
                                    maximumOffset + gLabel.PositionAndShape.BorderMarginBottom +
                                    this.rules.ChordSymbolBottomMargin); // TODO somehow off without margin for I numeral
                            } else {
                                gLabel.PositionAndShape.RelativePosition.y = minimumOffset + yShift;
                                gLabel.setLabelPositionAndShapeBorders();
                                gLabel.PositionAndShape.calculateBoundingBox();
                                skybottomcalculator.updateSkyLineInRange(start, end, minimumOffset + gLabel.PositionAndShape.BorderMarginTop);
                            }
                            previousChordContainer = graphicalChordContainer;
                        }
                    }
                }
            }
        }
    }

    protected calculateAlignedChordSymbolsOffset(staffEntries: GraphicalStaffEntry[], sbc: SkyBottomLineCalculator):
        {minOffset: number, maxOffset: number}
    {
        let minOffset: number = Number.MAX_SAFE_INTEGER;
        let maxOffset: number = Number.MIN_SAFE_INTEGER;
        for (const staffEntry of staffEntries) {
            for (const graphicalChordContainer of staffEntry.graphicalChordContainers) {
                const gps: BoundingBox = graphicalChordContainer.PositionAndShape;
                const parentBbox: BoundingBox = gps.Parent; // usually the staffEntry (bbox), but sometimes measure (for whole measure rests)
                let start: number = gps.BorderMarginLeft + parentBbox.AbsolutePosition.x;
                let end: number = gps.BorderMarginRight + parentBbox.AbsolutePosition.x;
                if (parentBbox.DataObject instanceof GraphicalMeasure) {
                    start += (parentBbox.DataObject as GraphicalMeasure).beginInstructionsWidth;
                    end += (parentBbox.DataObject as GraphicalMeasure).beginInstructionsWidth;
                }
                const placement: PlacementEnum = graphicalChordContainer.GetChordSymbolContainer.Placement;
                if (placement === PlacementEnum.Above) {
                    minOffset = Math.min(minOffset, sbc.getSkyLineMinInRange(start, end));
                } else if (placement === PlacementEnum.Below) {
                    maxOffset = Math.max(maxOffset, sbc.getBottomLineMaxInRange(start, end));
                }
            }
        }
        return {minOffset, maxOffset};
    }

    /**
     * Do layout on staff measures which only consist of a full rest.
     * @param rest
     * @param gse
     * @param measure
     */
    protected layoutMeasureWithWholeRest(rest: GraphicalNote, gse: GraphicalStaffEntry,
                                         measure: GraphicalMeasure): void {
        return;
    }

    protected layoutBeams(staffEntry: GraphicalStaffEntry): void {
        return;
    }

    protected layoutArticulationMarks(articulations: Articulation[], voiceEntry: VoiceEntry, graphicalStaffEntry: GraphicalStaffEntry): void {
        return;
    }

    protected layoutOrnament(ornaments: OrnamentContainer, voiceEntry: VoiceEntry,
                             graphicalStaffEntry: GraphicalStaffEntry): void {
        return;
    }

    protected calculateRestNotePlacementWithinGraphicalBeam(graphicalStaffEntry: GraphicalStaffEntry,
                                                            restNote: GraphicalNote,
                                                            previousNote: GraphicalNote,
                                                            nextStaffEntry: GraphicalStaffEntry,
                                                            nextNote: GraphicalNote): void {
        return;
    }

    protected calculateTupletNumbers(): void {
        if (!this.rules.TupletNumberLimitConsecutiveRepetitions) {
            return;
        }
        let currentTupletNumber: number = -1;
        let currentTypeLength: Fraction = undefined;
        let consecutiveTupletCount: number = 0;
        let currentTuplet: Tuplet = undefined;
        let skipTuplet: Tuplet = undefined; // if set, ignore (further) handling of this tuplet
        const disabledPerVoice: Object = {};
        for (const instrument of this.graphicalMusicSheet.ParentMusicSheet.Instruments) {
            for (const voice of instrument.Voices) {
                consecutiveTupletCount = 0; // reset for next voice
                disabledPerVoice[voice.VoiceId] = {};
                for (const ve of voice.VoiceEntries) {
                    if (ve.Notes.length > 0) {
                        const firstNote: Note = ve.Notes[0];
                        if (!firstNote.NoteTuplet ||
                            firstNote.NoteTuplet.shouldBeBracketed(
                                this.rules.TupletsBracketedUseXMLValue,
                                this.rules.TupletsBracketed,
                                this.rules.TripletsBracketed
                            )
                        ) {
                            // don't disable tuplet numbers under these conditions, reset consecutive tuplet count
                            currentTupletNumber = -1;
                            consecutiveTupletCount = 0;
                            currentTuplet = undefined;
                            currentTypeLength = undefined;
                            continue;
                        }
                        if (firstNote.NoteTuplet === skipTuplet) {
                            continue;
                        }
                        let typeLength: Fraction = firstNote.TypeLength;
                        if (!typeLength) {
                            // shouldn't happen, now that rest notes have TypeLength set too, see VoiceGenerator.addRestNote(), addSingleNote()
                            //   see test_tuplets_starting_with_rests_layout.mxl (first measure bass)
                            log.warn("note missing TypeLength");
                            typeLength = firstNote.NoteTuplet.Fractions[0];
                        }
                        if (firstNote.NoteTuplet !== currentTuplet) {
                            if (disabledPerVoice[voice.VoiceId][firstNote.NoteTuplet.TupletLabelNumber]) {
                                if (disabledPerVoice[voice.VoiceId][firstNote.NoteTuplet.TupletLabelNumber][typeLength.RealValue]) {
                                    firstNote.NoteTuplet.RenderTupletNumber = false;
                                    skipTuplet = firstNote.NoteTuplet;
                                    continue;
                                }
                            }
                        }
                        if (firstNote.NoteTuplet.TupletLabelNumber !== currentTupletNumber ||
                            !typeLength.Equals(currentTypeLength) ||
                            firstNote.NoteTuplet.Bracket) {
                            currentTupletNumber = firstNote.NoteTuplet.TupletLabelNumber;
                            currentTypeLength = typeLength;
                            consecutiveTupletCount = 0;
                        }
                        currentTuplet = firstNote.NoteTuplet;
                        consecutiveTupletCount++;
                        if (consecutiveTupletCount <= this.rules.TupletNumberMaxConsecutiveRepetitions) {
                            firstNote.NoteTuplet.RenderTupletNumber = true; // need to re-activate after re-render when it was set to false
                        }
                        if (consecutiveTupletCount > this.rules.TupletNumberMaxConsecutiveRepetitions) {
                            firstNote.NoteTuplet.RenderTupletNumber = false;
                            if (this.rules.TupletNumberAlwaysDisableAfterFirstMax) {
                                if (!disabledPerVoice[voice.VoiceId][currentTupletNumber]) {
                                    disabledPerVoice[voice.VoiceId][currentTupletNumber] = {};
                                }
                                disabledPerVoice[voice.VoiceId][currentTupletNumber][typeLength.RealValue] = true;
                            }
                        }
                        skipTuplet = currentTuplet;
                    }
                }
            }
        }
        return;
    }

    protected calculateSlurs(): void {
        return;
    }

    protected calculateGlissandi(): void {
        return;
    }

    protected calculateDynamicExpressionsForMultiExpression(multiExpression: MultiExpression, measureIndex: number, staffIndex: number): void {
        return;
    }


    /**
     * This method calculates the RelativePosition of a single verbal GraphicalContinuousDynamic.
     * @param graphicalContinuousDynamic Graphical continous dynamic to be calculated
     * @param startPosInStaffline Starting point in staff line
     */
    protected calculateGraphicalVerbalContinuousDynamic(graphicalContinuousDynamic: GraphicalContinuousDynamicExpression,
                                                        startPosInStaffline: PointF2D): void {
        // if ContinuousDynamicExpression is given from words
        const graphLabel: GraphicalLabel = graphicalContinuousDynamic.Label;
        const left: number = startPosInStaffline.x + graphLabel.PositionAndShape.BorderMarginLeft;
        const right: number = startPosInStaffline.x + graphLabel.PositionAndShape.BorderMarginRight;
        // placement always below the currentStaffLine, with the exception of Voice Instrument (-> above)
        const placement: PlacementEnum = graphicalContinuousDynamic.ContinuousDynamic.Placement;
        const staffLine: StaffLine = graphicalContinuousDynamic.ParentStaffLine;
        const skyBottomLineCalculator: SkyBottomLineCalculator = staffLine.SkyBottomLineCalculator;

        let drawingHeight: number;
        if (placement === PlacementEnum.Below) {
            drawingHeight = skyBottomLineCalculator.getBottomLineMaxInRange(left, right);    // Bottom line
            graphLabel.PositionAndShape.RelativePosition = new PointF2D(startPosInStaffline.x, drawingHeight - graphLabel.PositionAndShape.BorderMarginTop);
        } else {
            drawingHeight = skyBottomLineCalculator.getSkyLineMinInRange(left, right);
            graphLabel.PositionAndShape.RelativePosition = new PointF2D(startPosInStaffline.x, drawingHeight - graphLabel.PositionAndShape.BorderMarginBottom);
        }
    }

   /**
    * This method calculates the RelativePosition of a single GraphicalContinuousDynamic.
    * @param graphicalContinuousDynamic Graphical continous dynamic to be calculated
    * @param startPosInStaffline Starting point in staff line
    */
    public calculateGraphicalContinuousDynamic(graphicalContinuousDynamic: GraphicalContinuousDynamicExpression, startPosInStaffline: PointF2D): void {
        const isSoftAccent: boolean = graphicalContinuousDynamic.IsSoftAccent;
        const staffIndex: number = graphicalContinuousDynamic.ParentStaffLine.ParentStaff.idInMusicSheet;
        // TODO: Previously the staffIndex was passed down. BUT you can (and this function actually does this) get it from
        // the musicSystem OR from the ParentStaffLine. Is this the same index?
        // const staffIndex: number = musicSystem.StaffLines.indexOf(staffLine);

        // We know we have an end measure because otherwise we won't be called
        const endMeasure: GraphicalMeasure = this.graphicalMusicSheet.getGraphicalMeasureFromSourceMeasureAndIndex(
            graphicalContinuousDynamic.ContinuousDynamic.EndMultiExpression.SourceMeasureParent, staffIndex);
        if (!endMeasure) {
            log.warn("MusicSheetCalculator.calculateGraphicalContinuousDynamic: No endMeasure found");
            return;
        }

        graphicalContinuousDynamic.EndMeasure = endMeasure;
        const staffLine: StaffLine = graphicalContinuousDynamic.ParentStaffLine;
        const endStaffLine: StaffLine = endMeasure.ParentStaffLine;

        // check if Expression spreads over the same StaffLine or not
        const sameStaffLine: boolean = endStaffLine && staffLine === endStaffLine;

        let isPartOfMultiStaffInstrument: boolean = false;
        if (endStaffLine) { // unfortunately we can't do something like (endStaffLine?.check() || staffLine?.check()) in this typescript version
            isPartOfMultiStaffInstrument = endStaffLine?.isPartOfMultiStaffInstrument();
        } else if (staffLine) {
            isPartOfMultiStaffInstrument = staffLine?.isPartOfMultiStaffInstrument();
        }

        const endAbsoluteTimestamp: Fraction = Fraction.createFromFraction(graphicalContinuousDynamic.ContinuousDynamic.EndMultiExpression.AbsoluteTimestamp);
        const container: VerticalGraphicalStaffEntryContainer = this.graphicalMusicSheet.GetVerticalContainerFromTimestamp(endAbsoluteTimestamp);
        const parentMeasure: GraphicalMeasure = container.getFirstNonNullStaffEntry().parentMeasure;
        const endOfMeasure: number = parentMeasure.PositionAndShape.AbsolutePosition.x + parentMeasure.PositionAndShape.BorderRight;
        let maxNoteLength: Fraction = new Fraction(0, 0, 0);
        for (const staffEntry of container.StaffEntries) {
            if (staffEntry?.sourceStaffEntry.ParentStaff !== staffLine.ParentStaff) {
                // note: null check handles rare cases of undefined staffEntries, e.g. in test_wedge_cresc_dim_simultaneous_quartet.musicxml
                continue;
                // don't let notes in other staffs (not the wedge's staff) affect the wedge length (see #1477)
            }
            const currentMaxLength: Fraction = staffEntry?.sourceStaffEntry?.calculateMaxNoteLength(false);
            if ( currentMaxLength?.gt(maxNoteLength) ) {
                maxNoteLength = currentMaxLength;
            }
        }
        const useStaffEntryBorderLeft: boolean = !isSoftAccent &&
            graphicalContinuousDynamic.ContinuousDynamic.DynamicType === ContDynamicEnum.diminuendo;
        const endPosInStaffLine: PointF2D = this.getRelativePositionInStaffLineFromTimestamp(
            endAbsoluteTimestamp, staffIndex, endStaffLine, isPartOfMultiStaffInstrument, 0,
            useStaffEntryBorderLeft);

        const beginOfNextNote: Fraction = Fraction.plus(endAbsoluteTimestamp, maxNoteLength);
        const placementFraction: Fraction = beginOfNextNote.clone();
        const endOffsetFraction: Fraction = graphicalContinuousDynamic.ContinuousDynamic.EndMultiExpression.EndOffsetFraction;
        if (endOffsetFraction && this.rules.UseEndOffsetForExpressions) {
            placementFraction.Add(graphicalContinuousDynamic.ContinuousDynamic.EndMultiExpression.EndOffsetFraction);
        }
        // TODO for the last note of the piece (wedge ending after last note), this timestamp is incorrect, being after the last note
        //   but there's a workaround in getRelativePositionInStaffLineFromTimestamp() via the variable endAfterRightStaffEntry
        const nextNotePosInStaffLine: PointF2D = this.getRelativePositionInStaffLineFromTimestamp(
            placementFraction, staffIndex, endStaffLine, isPartOfMultiStaffInstrument, 0,
            graphicalContinuousDynamic.ContinuousDynamic.DynamicType === ContDynamicEnum.diminuendo);
        const wedgePadding: number = this.rules.SoftAccentWedgePadding;
        const staffEntryWidth: number = container.getFirstNonNullStaffEntry().PositionAndShape.Size.width; // staff entry widths for whole notes is too long
        const sizeFactor: number = this.rules.SoftAccentSizeFactor;
        //const standardWidth: number = 2;

        //If the next note position is not on the next staffline
        //extend close to the next note
        if (isSoftAccent) {
            //startPosInStaffline.x -= 1;
            startPosInStaffline.x -= staffEntryWidth / 2 * sizeFactor + wedgePadding;
            endPosInStaffLine.x = startPosInStaffline.x + staffEntryWidth / 2 * sizeFactor;
        } else if (nextNotePosInStaffLine.x > endPosInStaffLine.x && nextNotePosInStaffLine.x < endOfMeasure) {
            endPosInStaffLine.x += (nextNotePosInStaffLine.x - endPosInStaffLine.x) / this.rules.WedgeEndDistanceBetweenTimestampsFactor;
        } else { //Otherwise extend to the end of the measure
            endPosInStaffLine.x = endOfMeasure - this.rules.WedgeHorizontalMargin;
        }

        const startCollideBox: BoundingBox =
            this.dynamicExpressionMap.get(graphicalContinuousDynamic.ContinuousDynamic.StartMultiExpression.AbsoluteTimestamp.RealValue);
        if (startCollideBox) {
            if ((startCollideBox.DataObject as any).ParentStaffLine === staffLine) {
                // TODO the dynamicExpressionMap doesn't distinguish between staffLines, so we may react to a different staffline otherwise
                //   so the more fundamental solution would be to fix dynamicExpressionMap mapping across stafflines.
                startPosInStaffline.x = startCollideBox.RelativePosition.x + this.rules.WedgeHorizontalMargin;
                startPosInStaffline.x += startCollideBox.BorderMarginRight;
            }
        }
        //currentMusicSystem and currentStaffLine
        const musicSystem: MusicSystem = staffLine.ParentMusicSystem;
        const currentStaffLineIndex: number = musicSystem.StaffLines.indexOf(staffLine);
        const skyBottomLineCalculator: SkyBottomLineCalculator = staffLine.SkyBottomLineCalculator;
        // let expressionIndex: number;

        // placement always below the currentStaffLine, with the exception of Voice Instrument (-> above)
        const placement: PlacementEnum = graphicalContinuousDynamic.ContinuousDynamic.Placement;

        // if ContinuousDynamicExpression is given from wedge
        let endGraphicalContinuousDynamic: GraphicalContinuousDynamicExpression = undefined;

        // last length check
        if (sameStaffLine && endPosInStaffLine.x - startPosInStaffline.x < this.rules.WedgeMinLength && !isSoftAccent) {
            endPosInStaffLine.x = startPosInStaffline.x + this.rules.WedgeMinLength;
        }

        // First staff wedge always starts at the given position and the last and inbetween wedges always start at the begin of measure
        //   TODO: rename upper / lower to first / last, now that we can have inbetween wedges, though this creates a huge diff, and this should be clear now.
        const upperStartX: number = startPosInStaffline.x;
        let lowerStartX: number = endStaffLine.Measures[0].beginInstructionsWidth - this.rules.WedgeHorizontalMargin - 2;
        //TODO fix this when a range of measures to draw is given that doesn't include all the dynamic's measures (e.g. for crescendo)
        let upperEndX: number = 0;
        let lowerEndX: number = 0;

        /** Wedges between first and last staffline, in case we span more than 2 stafflines. */
        const inbetweenWedges: GraphicalContinuousDynamicExpression[] = [];
        if (!sameStaffLine) {
            // add wedge in all stafflines between (including) start and end measure
            upperEndX = staffLine.PositionAndShape.Size.width;
            lowerEndX = endPosInStaffLine.x;

            // get all stafflines between start measure and end measure, and add wedges for them.
            //   This would be less lines of code if there was already a list of stafflines for the sheet.
            const stafflinesCovered: StaffLine[] = [staffLine, endStaffLine]; // start and end staffline already get a wedge
            const startMeasure: GraphicalMeasure = graphicalContinuousDynamic.StartMeasure;
            let nextMeasure: GraphicalMeasure = startMeasure;
            let iterations: number = 0; // safety measure against infinite loop
            let sourceMeasureIndex: number = startMeasure.parentSourceMeasure.measureListIndex;
            while (nextMeasure !== endMeasure && iterations < 1000) {
                const nextSourceMeasure: SourceMeasure = this.graphicalMusicSheet.ParentMusicSheet.SourceMeasures[sourceMeasureIndex];
                const potentialNextMeasure: GraphicalMeasure = this.graphicalMusicSheet.getGraphicalMeasureFromSourceMeasureAndIndex(
                    nextSourceMeasure, staffIndex
                );
                if (potentialNextMeasure) {
                    nextMeasure = potentialNextMeasure;
                    const nextStaffline: StaffLine = nextMeasure.ParentStaffLine;
                    if (!stafflinesCovered.includes(nextStaffline)) {
                        stafflinesCovered.push(nextStaffline);
                        const newWedge: GraphicalContinuousDynamicExpression =
                            new GraphicalContinuousDynamicExpression(
                                graphicalContinuousDynamic.ContinuousDynamic,
                                nextStaffline,
                                nextStaffline.Measures[0].parentSourceMeasure
                            );
                        newWedge.IsSplittedPart = true;
                        inbetweenWedges.push(newWedge);
                    }
                }
                sourceMeasureIndex++;
                iterations++;
            }

            // last wedge at endMeasure
            endGraphicalContinuousDynamic = new GraphicalContinuousDynamicExpression(
                graphicalContinuousDynamic.ContinuousDynamic, endStaffLine, endMeasure.parentSourceMeasure);
            endGraphicalContinuousDynamic.IsSplittedPart = true;
            graphicalContinuousDynamic.IsSplittedPart = true;
        } else {
            upperEndX = endPosInStaffLine.x;
        }
        if (isSoftAccent) {
            // secondGraphicalContinuousDynamic = new GraphicalContinuousDynamicExpression(
            //     graphicalContinuousDynamic.ContinuousDynamic,
            //     graphicalContinuousDynamic.ParentStaffLine,
            //     graphicalContinuousDynamic.StartMeasure.parentSourceMeasure
            // );
            // secondGraphicalContinuousDynamic.StartIsEnd = true;
            // doesn't work well with secondGraphicalDynamic, positions/rendering messed up
            lowerStartX = endPosInStaffLine.x + wedgePadding;
            lowerEndX = lowerStartX + staffEntryWidth / 2 * sizeFactor;
        }

        // the Height of the Expression's placement
        let idealY: number = 0;
        let endIdealY: number = 0;

        if (placement === PlacementEnum.Below) {
            // can be a single Staff Instrument or an Instrument with 2 Staves
            let nextStaffLineIndex: number = 0;
            if (currentStaffLineIndex < musicSystem.StaffLines.length - 1) {
                nextStaffLineIndex = currentStaffLineIndex + 1;
            }

            // check, maybe currentStaffLine is the last of the MusicSystem (and it has a ContinuousDynamicExpression with placement below)
            if (nextStaffLineIndex > currentStaffLineIndex) {
                // currentStaffLine isn't the last of the MusicSystem
                const nextStaffLine: StaffLine = musicSystem.StaffLines[nextStaffLineIndex];

                const distanceBetweenStaffLines: number = nextStaffLine.PositionAndShape.RelativePosition.y -
                    staffLine.PositionAndShape.RelativePosition.y -
                    this.rules.StaffHeight;

                // ideal Height is exactly between the two StaffLines
                idealY = this.rules.StaffHeight + distanceBetweenStaffLines / 2;
            } else {
                // currentStaffLine is the MusicSystem's last
                idealY = this.rules.WedgePlacementBelowY;
            }

            // must consider the upperWedge starting/ending tip for the comparison with the BottomLine
            idealY -= this.rules.WedgeOpeningLength / 2;
            if (!sameStaffLine) {
                // Set the value for the splitted y position to the ideal position before we check and modify it with
                // the skybottom calculator detection
                endIdealY = idealY;
            }
            // must check BottomLine for possible collisions within the Length of the Expression
            // find the corresponding max value for the given Length
            let maxBottomLineValueForExpressionLength: number = skyBottomLineCalculator.getBottomLineMaxInRange(upperStartX, upperEndX);

            // if collisions, then set the Height accordingly
            if (maxBottomLineValueForExpressionLength > idealY) {
                idealY = maxBottomLineValueForExpressionLength;
            }

            // special case - wedge must be drawn within the boundaries of a crossedBeam
            const withinCrossedBeam: boolean = false;

            if (currentStaffLineIndex < musicSystem.StaffLines.length - 1) {
                // find GraphicalStaffEntries closest to wedge's xPositions
                const closestToEndStaffEntry: GraphicalStaffEntry = staffLine.findClosestStaffEntry(upperEndX);
                const closestToStartStaffEntry: GraphicalStaffEntry = staffLine.findClosestStaffEntry(upperStartX);

                if (closestToStartStaffEntry && closestToEndStaffEntry) {
                    // must check both StaffLines
                    const startVerticalContainer: VerticalGraphicalStaffEntryContainer = closestToStartStaffEntry.parentVerticalContainer;
                    // const endVerticalContainer: VerticalGraphicalStaffEntryContainer = closestToEndStaffEntry.parentVerticalContainer;
                    if (startVerticalContainer) {
                        // TODO: Needs to be implemented?
                        // withinCrossedBeam = areStaffEntriesWithinCrossedBeam(startVerticalContainer,
                        // endVerticalContainer, currentStaffLineIndex, nextStaffLineIndex);
                    }

                    if (withinCrossedBeam) {
                        const nextStaffLine: StaffLine = musicSystem.StaffLines[nextStaffLineIndex];
                        const nextStaffLineMinSkyLineValue: number = nextStaffLine.SkyBottomLineCalculator.getSkyLineMinInRange(upperStartX, upperEndX);
                        const distanceBetweenStaffLines: number = nextStaffLine.PositionAndShape.RelativePosition.y -
                            staffLine.PositionAndShape.RelativePosition.y;
                        const relativeSkyLineHeight: number = distanceBetweenStaffLines + nextStaffLineMinSkyLineValue;

                        if (relativeSkyLineHeight - this.rules.WedgeOpeningLength > this.rules.StaffHeight) {
                            idealY = relativeSkyLineHeight - this.rules.WedgeVerticalMargin;
                        } else {
                            idealY = this.rules.StaffHeight + this.rules.WedgeOpeningLength;
                        }

                        graphicalContinuousDynamic.NotToBeRemoved = true;
                    }
                }
            }

            // do the same in case of a Wedge ending at another StaffLine
            if (!sameStaffLine) {
                maxBottomLineValueForExpressionLength = endStaffLine.SkyBottomLineCalculator.getBottomLineMaxInRange(lowerStartX, lowerEndX);

                if (maxBottomLineValueForExpressionLength > endIdealY) {
                    endIdealY = maxBottomLineValueForExpressionLength;
                }

                endIdealY += this.rules.WedgeOpeningLength / 2;
                endIdealY += this.rules.WedgeVerticalMargin;
            }

            if (!withinCrossedBeam) {
                idealY += this.rules.WedgeOpeningLength / 2;
                idealY += this.rules.WedgeVerticalMargin;
            }

        } else if (placement === PlacementEnum.Above) {
            // single Staff Instrument (eg Voice)
            if (staffLine.ParentStaff.ParentInstrument.Staves.length === 1) {
                // single Staff Voice Instrument
                idealY = this.rules.WedgePlacementAboveY;
            } else {
                // Staff = not the first Staff of a 2-staved Instrument
                let previousStaffLineIndex: number = 0;
                if (currentStaffLineIndex > 0) {
                    previousStaffLineIndex = currentStaffLineIndex - 1;
                }

                const previousStaffLine: StaffLine = musicSystem.StaffLines[previousStaffLineIndex];
                const distanceBetweenStaffLines: number = staffLine.PositionAndShape.RelativePosition.y -
                    previousStaffLine.PositionAndShape.RelativePosition.y -
                    this.rules.StaffHeight;

                // ideal Height is exactly between the two StaffLines
                idealY = -distanceBetweenStaffLines / 2;
            }

            // must consider the upperWedge starting/ending tip for the comparison with the SkyLine
            idealY += this.rules.WedgeOpeningLength / 2;
            if (!sameStaffLine) {
                endIdealY = idealY;
            }

            // must check SkyLine for possible collisions within the Length of the Expression
            // find the corresponding min value for the given Length
            let minSkyLineValueForExpressionLength: number = skyBottomLineCalculator.getSkyLineMinInRange(upperStartX, upperEndX);

            // if collisions, then set the Height accordingly
            if (minSkyLineValueForExpressionLength < idealY) {
                idealY = minSkyLineValueForExpressionLength;
            }
            const withinCrossedBeam: boolean = false;

            // special case - wedge must be drawn within the boundaries of a crossedBeam
            if (staffLine.ParentStaff.ParentInstrument.Staves.length > 1 && currentStaffLineIndex > 0) {
                // find GraphicalStaffEntries closest to wedge's xPositions
                const closestToStartStaffEntry: GraphicalStaffEntry = staffLine.findClosestStaffEntry(upperStartX);
                const closestToEndStaffEntry: GraphicalStaffEntry = staffLine.findClosestStaffEntry(upperEndX);

                if (closestToStartStaffEntry && closestToEndStaffEntry) {
                    // must check both StaffLines
                    const startVerticalContainer: VerticalGraphicalStaffEntryContainer = closestToStartStaffEntry.parentVerticalContainer;
                    // const endVerticalContainer: VerticalGraphicalStaffEntryContainer = closestToEndStaffEntry.parentVerticalContainer;
                    const formerStaffLineIndex: number = currentStaffLineIndex - 1;
                    if (startVerticalContainer) {
                        // withinCrossedBeam = this.areStaffEntriesWithinCrossedBeam(startVerticalContainer,
                        // endVerticalContainer, currentStaffLineIndex, formerStaffLineIndex);
                    }

                    if (withinCrossedBeam) {
                        const formerStaffLine: StaffLine = musicSystem.StaffLines[formerStaffLineIndex];
                        if (formerStaffLine) { // can be undefined if staff.Visible = false (e.g. piano)
                            const formerStaffLineMaxBottomLineValue: number = formerStaffLine.SkyBottomLineCalculator.
                                                                              getBottomLineMaxInRange(upperStartX, upperEndX);
                            const distanceBetweenStaffLines: number = staffLine.PositionAndShape.RelativePosition.y -
                                formerStaffLine.PositionAndShape.RelativePosition.y;
                            const relativeSkyLineHeight: number = distanceBetweenStaffLines - formerStaffLineMaxBottomLineValue;
                            idealY = (relativeSkyLineHeight - this.rules.StaffHeight) / 2 + this.rules.StaffHeight;
                        }
                    }
                }
            }

            // do the same in case of a Wedge ending at another StaffLine
            if (!sameStaffLine) {
                minSkyLineValueForExpressionLength = endStaffLine.SkyBottomLineCalculator.getSkyLineMinInRange(lowerStartX, lowerEndX);

                if (minSkyLineValueForExpressionLength < endIdealY) {
                    endIdealY = minSkyLineValueForExpressionLength;
                }

                endIdealY -= this.rules.WedgeOpeningLength / 2;
            }

            if (!withinCrossedBeam) {
                idealY -= this.rules.WedgeOpeningLength / 2;
                idealY -= this.rules.WedgeVerticalMargin;
            }
            if (!sameStaffLine) {
                endIdealY -= this.rules.WedgeVerticalMargin;
            }
        }

        // now we have the correct placement Height for the Expression
        // the idealY is calculated relative to the currentStaffLine

        graphicalContinuousDynamic.Lines.clear();
        // create wedges (crescendo / decrescendo lines)
        if (isSoftAccent) {
            graphicalContinuousDynamic.createFirstHalfCrescendoLines(upperStartX, upperEndX, idealY);
            graphicalContinuousDynamic.createSecondHalfDiminuendoLines(lowerStartX, lowerEndX, idealY);
            graphicalContinuousDynamic.calcPsi();
        } else if (sameStaffLine && !isSoftAccent) {
            // either create crescendo or decrescendo lines, same principle / parameters.
            graphicalContinuousDynamic.createLines(upperStartX, upperEndX, idealY);
            graphicalContinuousDynamic.calcPsi();
        } else {
            // two+ different Wedges
            // first wedge
            graphicalContinuousDynamic.createFirstHalfLines(upperStartX, upperEndX, idealY);
            graphicalContinuousDynamic.calcPsi();

            // inbetween wedges
            for (let i: number = 0; i < inbetweenWedges.length; i++) {
                const inbetweenWedge: GraphicalContinuousDynamicExpression = inbetweenWedges[i];
                const inbetweenStaffline: StaffLine = inbetweenWedge.ParentStaffLine;
                let betweenIdealY: number = endIdealY;

                if (placement === PlacementEnum.Below) {
                    const maxBottomLineValueForExpressionLength: number =
                    endStaffLine.SkyBottomLineCalculator.getBottomLineMaxInRange(lowerStartX, upperEndX);
                    if (maxBottomLineValueForExpressionLength > betweenIdealY) {
                        betweenIdealY = maxBottomLineValueForExpressionLength;
                    }
                    betweenIdealY += this.rules.WedgeOpeningLength / 2;
                    betweenIdealY += this.rules.WedgeVerticalMargin;
                } else if (placement === PlacementEnum.Above) {
                    const minSkyLineValueForExpressionLength: number =
                        inbetweenStaffline.SkyBottomLineCalculator.getSkyLineMinInRange(lowerStartX, lowerEndX);
                    if (minSkyLineValueForExpressionLength < endIdealY) {
                        betweenIdealY = minSkyLineValueForExpressionLength;
                    }
                    betweenIdealY -= this.rules.WedgeOpeningLength / 2;
                }

                if (graphicalContinuousDynamic.ContinuousDynamic.DynamicType === ContDynamicEnum.crescendo) {
                    inbetweenWedge.createSecondHalfCrescendoLines(0, inbetweenStaffline.PositionAndShape.Size.width, betweenIdealY);
                    // for crescendo, we want the same look as on the last staffline: not starting with an intersection / starting wedge
                } else {
                    inbetweenWedge.createFirstHalfDiminuendoLines(0, inbetweenStaffline.PositionAndShape.Size.width, betweenIdealY);
                    // for diminuendo, we want the same look as on the first staffline: not ending in an intersection / looking finished
                }
                inbetweenWedge.calcPsi();
            }

            // last wedge
            endGraphicalContinuousDynamic.createSecondHalfLines(lowerStartX, lowerEndX, endIdealY);
            endGraphicalContinuousDynamic.calcPsi();
        }
        this.dynamicExpressionMap.set(endAbsoluteTimestamp.RealValue, graphicalContinuousDynamic.PositionAndShape);
    }

    /**
     * This method calculates the RelativePosition of a single GraphicalInstantaneousDynamicExpression.
     * @param graphicalInstantaneousDynamic Dynamic expression to be calculated
     * @param startPosInStaffline Starting point in staff line
     */
    protected calculateGraphicalInstantaneousDynamicExpression(graphicalInstantaneousDynamic: GraphicalInstantaneousDynamicExpression,
                                                               startPosInStaffline: PointF2D, timestamp: Fraction): void {
        // get Margin Dimensions
        const staffLine: StaffLine = graphicalInstantaneousDynamic.ParentStaffLine;
        if (!staffLine) {
            return; // TODO can happen when drawing range modified (osmd.setOptions({drawFromMeasureNumber...}))
        }

        const left: number = startPosInStaffline.x + graphicalInstantaneousDynamic.PositionAndShape.BorderMarginLeft;
        const right: number = startPosInStaffline.x + graphicalInstantaneousDynamic.PositionAndShape.BorderMarginRight;
        const skyBottomLineCalculator: SkyBottomLineCalculator = staffLine.SkyBottomLineCalculator;
        let yPosition: number = 0;

        // calculate yPosition according to Placement
        if (graphicalInstantaneousDynamic.Placement === PlacementEnum.Above) {
            const skyLineValue: number = skyBottomLineCalculator.getSkyLineMinInRange(left, right);

            // if StaffLine part of multiStaff Instrument and not the first one, ideal yPosition middle of distance between Staves
            if (staffLine.isPartOfMultiStaffInstrument() && staffLine.ParentStaff !== staffLine.ParentStaff.ParentInstrument.Staves[0]) {
                const formerStaffLine: StaffLine = staffLine.ParentMusicSystem.StaffLines[staffLine.ParentMusicSystem.StaffLines.indexOf(staffLine) - 1];
                if (formerStaffLine) { // can be undefined if staff.Visible = false (e.g. piano)
                    const difference: number = staffLine?.PositionAndShape.RelativePosition.y -
                        formerStaffLine.PositionAndShape.RelativePosition.y - this.rules.StaffHeight;
                    // take always into account the size of the Dynamic
                    if (skyLineValue > -difference / 2) {
                        yPosition = -difference / 2;
                    } else {
                        yPosition = skyLineValue - graphicalInstantaneousDynamic.PositionAndShape.BorderMarginBottom;
                    }
                }
            } else {
                yPosition = skyLineValue - graphicalInstantaneousDynamic.PositionAndShape.BorderMarginBottom;
            }

            graphicalInstantaneousDynamic.PositionAndShape.RelativePosition = new PointF2D(startPosInStaffline.x, yPosition);
        } else if (graphicalInstantaneousDynamic.Placement === PlacementEnum.Below) {
            const bottomLineValue: number = skyBottomLineCalculator.getBottomLineMaxInRange(left, right);
            // if StaffLine part of multiStaff Instrument and not the last one, ideal yPosition middle of distance between Staves
            const lastStaff: Staff = staffLine.ParentStaff.ParentInstrument.Staves[staffLine.ParentStaff.ParentInstrument.Staves.length - 1];
            if (staffLine.isPartOfMultiStaffInstrument() && staffLine.ParentStaff !== lastStaff) {
                const nextStaffLine: StaffLine = staffLine.ParentMusicSystem.StaffLines[staffLine.ParentMusicSystem.StaffLines.indexOf(staffLine) + 1];
                if (nextStaffLine) {
                    // nextStaffLine can be undefined if one staff of an instrument (e.g. piano left hand) is invisible (Visible = false)
                    const difference: number = nextStaffLine.PositionAndShape.RelativePosition.y -
                        staffLine.PositionAndShape.RelativePosition.y - this.rules.StaffHeight;
                    const border: number = graphicalInstantaneousDynamic.PositionAndShape.BorderMarginBottom;
                    // take always into account the size of the Dynamic
                    if (bottomLineValue + border < this.rules.StaffHeight + difference / 2) {
                        yPosition = this.rules.StaffHeight + difference / 2;
                    } else {
                        yPosition = bottomLineValue - graphicalInstantaneousDynamic.PositionAndShape.BorderMarginTop;
                    }
                }
            } else {
                yPosition = bottomLineValue - graphicalInstantaneousDynamic.PositionAndShape.BorderMarginTop;
            }

            graphicalInstantaneousDynamic.PositionAndShape.RelativePosition = new PointF2D(startPosInStaffline.x, yPosition);
        }
        graphicalInstantaneousDynamic.updateSkyBottomLine();
    }

    protected calcGraphicalRepetitionEndingsRecursively(repetition: Repetition): void {
        return;
    }

    /**
     * Calculate a single GraphicalRepetition.
     * @param start
     * @param end
     * @param numberText
     * @param offset
     * @param leftOpen
     * @param rightOpen
     */
    protected layoutSingleRepetitionEnding(start: GraphicalMeasure, end: GraphicalMeasure, numberText: string,
                                           offset: number, leftOpen: boolean, rightOpen: boolean): void {
        return;
    }

    protected calculateLabel(staffLine: StaffLine,
                             relative: PointF2D,
                             combinedString: string,
                             style: FontStyles,
                             placement: PlacementEnum,
                             fontHeight: number,
                             textAlignment: TextAlignmentEnum = TextAlignmentEnum.CenterBottom,
                             yPadding: number = 0): GraphicalLabel {
        const label: Label = new Label(combinedString, textAlignment);
        label.fontStyle = style;
        label.fontHeight = fontHeight;

        // TODO_RR: TextHeight from first Entry
        const graphLabel: GraphicalLabel = new GraphicalLabel(label, fontHeight, label.textAlignment, this.rules, staffLine.PositionAndShape);
        const marginFactor: number = 1.1;

        if (placement === PlacementEnum.Below) {
            graphLabel.Label.textAlignment = TextAlignmentEnum.LeftTop;
        }

        graphLabel.setLabelPositionAndShapeBorders();
        graphLabel.PositionAndShape.BorderMarginBottom *= marginFactor;
        graphLabel.PositionAndShape.BorderMarginTop *= marginFactor;
        graphLabel.PositionAndShape.BorderMarginLeft *= marginFactor;
        graphLabel.PositionAndShape.BorderMarginRight *= marginFactor;

        let left: number = relative.x + graphLabel.PositionAndShape.BorderMarginLeft;
        let right: number = relative.x + graphLabel.PositionAndShape.BorderMarginRight;

        // check if GraphicalLabel exceeds the StaffLine's borders.
        if (right > staffLine.PositionAndShape.Size.width) {
            right = staffLine.PositionAndShape.Size.width - this.rules.MeasureRightMargin;
            left = right - graphLabel.PositionAndShape.MarginSize.width;
            relative.x = left - graphLabel.PositionAndShape.BorderMarginLeft;
        }
        if (left < staffLine.PositionAndShape.BorderMarginLeft) {
            const rightShift: number = staffLine.PositionAndShape.BorderMarginLeft - left + this.rules.LabelXOffsetForStafflineLeftOverflowCheck;
            left += rightShift;
            right += rightShift;
            relative.x += rightShift;
        }

        // find allowed position (where the Label can be positioned) from Sky- BottomLine
        let drawingHeight: number;
        const skyBottomLineCalculator: SkyBottomLineCalculator = staffLine.SkyBottomLineCalculator;
        if (placement === PlacementEnum.Below) {
            drawingHeight = skyBottomLineCalculator.getBottomLineMaxInRange(left, right) + yPadding;
        } else {
            drawingHeight = skyBottomLineCalculator.getSkyLineMinInRange(left, right) - yPadding;
        }

        // set RelativePosition
        graphLabel.PositionAndShape.RelativePosition = new PointF2D(relative.x, drawingHeight);

        // update Sky- BottomLine
        if (placement === PlacementEnum.Below) {
            skyBottomLineCalculator.updateBottomLineInRange(left, right, graphLabel.PositionAndShape.BorderMarginBottom + drawingHeight);
        } else {
            skyBottomLineCalculator.updateSkyLineInRange(left, right, graphLabel.PositionAndShape.BorderMarginTop + drawingHeight);
        }
        return graphLabel;
    }

    protected calculateTempoExpressionsForMultiTempoExpression(sourceMeasure: SourceMeasure, multiTempoExpression: MultiTempoExpression,
                                                               measureIndex: number): void {
        // calculate absolute Timestamp
        const absoluteTimestamp: Fraction = Fraction.plus(sourceMeasure.AbsoluteTimestamp, multiTempoExpression.Timestamp);
        const measures: GraphicalMeasure[] = this.graphicalMusicSheet.MeasureList[measureIndex];
        let relative: PointF2D = new PointF2D();

        if (multiTempoExpression.ContinuousTempo || multiTempoExpression.InstantaneousTempo) {
            // TempoExpressions always on the first visible System's StaffLine // TODO is it though?
            if (this.rules.MinMeasureToDrawIndex > 0) {
                return; // assuming that the tempo is always in measure 1 (idx 0), adding the expression causes issues when we don't draw measure 1
            }
            if (!measures[0]) {
                return;
            }
            let staffLine: StaffLine = measures[0].ParentStaffLine;
            let firstVisibleMeasureX: number = measures[0].PositionAndShape.RelativePosition.x;
            let verticalIndex: number = 0;
            for (let j: number = 0; j < measures.length; j++) {
                if (!measures[j].ParentStaffLine || measures[j].ParentStaffLine.Measures.length === 0) {
                    continue;
                }

                if (measures[j].ParentStaffLine.Measures.length > 0) {
                    staffLine = measures[j].ParentStaffLine;
                    firstVisibleMeasureX = measures[j].PositionAndShape.RelativePosition.x;
                    verticalIndex = j;
                    break;
                }
            }
            relative = this.getRelativePositionInStaffLineFromTimestamp(absoluteTimestamp,
                                                                        verticalIndex,
                                                                        staffLine,
                                                                        staffLine.isPartOfMultiStaffInstrument(),
                                                                        firstVisibleMeasureX);

            // also placement Above
            if (multiTempoExpression.EntriesList.length > 0 &&
                multiTempoExpression.EntriesList[0].Expression instanceof InstantaneousTempoExpression) {
                const instantaniousTempo: InstantaneousTempoExpression = (multiTempoExpression.EntriesList[0].Expression as InstantaneousTempoExpression);
                instantaniousTempo.Placement = PlacementEnum.Above;

                // if an InstantaniousTempoExpression exists at the very beginning then
                // check if expression is positioned at first ever StaffEntry and
                // check if MusicSystem is first MusicSystem
                if (staffLine.Measures[0].staffEntries.length > 0 &&
                    Math.abs(relative.x - staffLine.Measures[0].staffEntries[0].PositionAndShape.RelativePosition.x) === 0 &&
                    staffLine.ParentMusicSystem === this.musicSystems[0]) {
                    const firstInstructionEntry: GraphicalStaffEntry = staffLine.Measures[0].FirstInstructionStaffEntry;
                    if (firstInstructionEntry) {
                        const lastInstruction: AbstractGraphicalInstruction = firstInstructionEntry.GraphicalInstructions.last();
                        relative.x = lastInstruction.PositionAndShape.RelativePosition.x;
                    }
                    if (this.rules.CompactMode) {
                        relative.x = staffLine.PositionAndShape.RelativePosition.x +
                            staffLine.Measures[0].PositionAndShape.RelativePosition.x;
                    }
                }
            }

            // const addAtLastList: GraphicalObject[] = [];
            for (const entry of multiTempoExpression.EntriesList) {
                let textAlignment: TextAlignmentEnum = this.rules.TempoExpressionTextAlignment;
                if (this.rules.CompactMode) {
                    textAlignment = TextAlignmentEnum.LeftBottom;
                }
                const graphLabel: GraphicalLabel = this.calculateLabel(staffLine,
                                                                       relative,
                                                                       entry.label,
                                                                       multiTempoExpression.getFontstyleOfFirstEntry(),
                                                                       entry.Expression.Placement,
                                                                       this.rules.UnknownTextHeight,
                                                                       textAlignment,
                                                                       this.rules.TempoYSpacing);
                if (entry.Expression.ColorXML && this.rules.ExpressionsUseXMLColor) {
                    graphLabel.ColorXML = entry.Expression.ColorXML;
                }

                if (entry.Expression instanceof InstantaneousTempoExpression) {
                    //already added?
                    for (const expr of staffLine.AbstractExpressions) {
                        if (expr instanceof GraphicalInstantaneousTempoExpression &&
                            (expr.SourceExpression as AbstractTempoExpression).Label === entry.Expression.Label) {
                            //already added
                            continue;
                        }
                    }

                    const graphicalTempoExpr: GraphicalInstantaneousTempoExpression = new GraphicalInstantaneousTempoExpression(entry.Expression, graphLabel);
                    if (!graphicalTempoExpr.ParentStaffLine) {
                        log.warn("Adding staffline didn't work");
                        // I am actually fooling the linter here and use the created object. This method needs refactoring,
                        // all graphical expression creations should be in one place and have basic stuff like labels, lines, ...
                        // in their constructor
                    }
                    // in case of metronome mark:
                    if (this.rules.MetronomeMarksDrawn) {
                        if ((entry.Expression as InstantaneousTempoExpression).TempoType === TempoType.metronomeMark) {
                            this.createMetronomeMark((entry.Expression as InstantaneousTempoExpression));
                            continue;
                        }
                    }
                } else if (entry.Expression instanceof ContinuousTempoExpression) {
                    for (const expr of staffLine.AbstractExpressions) {
                        if (expr instanceof GraphicalInstantaneousTempoExpression &&
                        (expr.SourceExpression as AbstractTempoExpression).Label === entry.Expression.Label) {
                            continue; // already added
                        }
                    }
                    // TODO maybe create GraphicalContinuousTempoExpression class,
                    //   though the ContinuousTempoExpressions we have currently behave the same graphically (accelerando, ritardando, etc).
                    //   The behavior difference rather affects playback (e.g. ritardando, which gradually changes tempo)
                    staffLine.AbstractExpressions.push(new GraphicalInstantaneousTempoExpression(entry.Expression, graphLabel));
                }
            }
        }
    }

    protected createMetronomeMark(metronomeExpression: InstantaneousTempoExpression): void {
        throw new Error(this.abstractNotImplementedErrorMessage);
    }

    protected graphicalMeasureCreatedCalculations(measure: GraphicalMeasure): void {
        return;
    }

    protected clearSystemsAndMeasures(): void {
        for (let idx: number = 0, len: number = this.graphicalMusicSheet.MusicPages.length; idx < len; ++idx) {
            const graphicalMusicPage: GraphicalMusicPage = this.graphicalMusicSheet.MusicPages[idx];
            for (let idx2: number = 0, len2: number = graphicalMusicPage.MusicSystems.length; idx2 < len2; ++idx2) {
                const musicSystem: MusicSystem = graphicalMusicPage.MusicSystems[idx2];
                for (let idx3: number = 0, len3: number = musicSystem.StaffLines.length; idx3 < len3; ++idx3) {
                    const staffLine: StaffLine = musicSystem.StaffLines[idx3];
                    if (!staffLine.ParentStaff.Visible) {
                        staffLine.Measures.clear();
                        // musicSystem.PositionAndShape.ChildElements = musicSystem.PositionAndShape.ChildElements.filter(
                        //     (child) => child !== staffLine.PositionAndShape
                        // );
                    }
                    // if (!staffLine) {
                    //     continue;
                    // }
                    // if (!staffLine.ParentStaff.Visible) {
                    //     musicSystem.StaffLines = musicSystem.StaffLines.slice(idx3);
                    //     musicSystem.PositionAndShape.ChildElements = musicSystem.PositionAndShape.ChildElements.filter(
                    //         (child) => child !== staffLine.PositionAndShape
                    //     );
                    //     musicSystem.PositionAndShape.calculateBoundingBox();
                    //     idx3--;
                    //     continue;
                    // }
                    for (let idx4: number = 0, len4: number = staffLine.Measures.length; idx4 < len4; ++idx4) {
                        const graphicalMeasure: GraphicalMeasure = staffLine.Measures[idx4];
                        if (graphicalMeasure.FirstInstructionStaffEntry) {
                            const index: number = graphicalMeasure.PositionAndShape.ChildElements.indexOf(
                                graphicalMeasure.FirstInstructionStaffEntry.PositionAndShape
                            );
                            if (index > -1) {
                                graphicalMeasure.PositionAndShape.ChildElements.splice(index, 1);
                            }
                            graphicalMeasure.FirstInstructionStaffEntry = undefined;
                            graphicalMeasure.beginInstructionsWidth = 0.0;
                        }
                        if (graphicalMeasure.LastInstructionStaffEntry) {
                            const index: number = graphicalMeasure.PositionAndShape.ChildElements.indexOf(
                                graphicalMeasure.LastInstructionStaffEntry.PositionAndShape
                            );
                            if (index > -1) {
                                graphicalMeasure.PositionAndShape.ChildElements.splice(index, 1);
                            }
                            graphicalMeasure.LastInstructionStaffEntry = undefined;
                            graphicalMeasure.endInstructionsWidth = 0.0;
                        }
                    }
                    staffLine.Measures = [];
                    staffLine.PositionAndShape.ChildElements = [];
                }
                musicSystem.StaffLines.length = 0;
                musicSystem.PositionAndShape.ChildElements = [];
            }
            graphicalMusicPage.MusicSystems = [];
            graphicalMusicPage.PositionAndShape.ChildElements = [];
        }
        this.graphicalMusicSheet.MusicPages = [];
    }

    protected handleVoiceEntry(voiceEntry: VoiceEntry, graphicalStaffEntry: GraphicalStaffEntry,
                               accidentalCalculator: AccidentalCalculator, openLyricWords: LyricWord[],
                               activeClef: ClefInstruction,
                               openTuplets: Tuplet[], openBeams: Beam[],
                               octaveShiftValue: OctaveEnum, staffIndex: number,
                               linkedNotes: Note[] = undefined,
                               sourceStaffEntry: SourceStaffEntry = undefined): OctaveEnum {
        if (voiceEntry.StemDirectionXml !== StemDirectionType.Undefined &&
            this.rules.SetWantedStemDirectionByXml &&
            voiceEntry.StemDirectionXml !== undefined) {
                voiceEntry.WantedStemDirection = voiceEntry.StemDirectionXml;
        } else {
            this.calculateStemDirectionFromVoices(voiceEntry);
        }
        // if GraphicalStaffEntry has been created earlier (because of Tie), then the GraphicalNotesLists have also been created
        const gve: GraphicalVoiceEntry = graphicalStaffEntry.findOrCreateGraphicalVoiceEntry(voiceEntry);
        gve.octaveShiftValue = octaveShiftValue;
        // check for Tabs:
        const tabStaffEntry: GraphicalStaffEntry = graphicalStaffEntry.tabStaffEntry;
        let graphicalTabVoiceEntry: GraphicalVoiceEntry;
        if (tabStaffEntry) {
            graphicalTabVoiceEntry = tabStaffEntry.findOrCreateGraphicalVoiceEntry(voiceEntry);
        }

        for (let idx: number = 0, len: number = voiceEntry.Notes.length; idx < len; ++idx) {
            const note: Note = voiceEntry.Notes[idx];
            if (!note) {
                continue;
            }
            if (sourceStaffEntry !== undefined && sourceStaffEntry.Link !== undefined && linkedNotes !== undefined && linkedNotes.indexOf(note) > -1) {
                continue;
            }
            let graphicalNote: GraphicalNote;
            if (voiceEntry.IsGrace) {
                graphicalNote = MusicSheetCalculator.symbolFactory.createGraceNote(note, gve, activeClef, this.rules, octaveShiftValue);
            } else {
                graphicalNote = MusicSheetCalculator.symbolFactory.createNote(note, gve, activeClef, octaveShiftValue, this.rules, undefined);
                MusicSheetCalculator.stafflineNoteCalculator.trackNote(graphicalNote);
            }
            if (note.Pitch) {
                this.checkNoteForAccidental(graphicalNote, accidentalCalculator, activeClef, octaveShiftValue);
            }
            this.resetYPositionForLeadSheet(graphicalNote.PositionAndShape);
            graphicalStaffEntry.addGraphicalNoteToListAtCorrectYPosition(gve, graphicalNote);
            graphicalNote.PositionAndShape.calculateBoundingBox();
            if (!this.leadSheet) {
                if (note.NoteBeam !== undefined && note.PrintObject) {
                    if (!(note instanceof TabNote) || this.rules.TabBeamsRendered) {
                        this.handleBeam(graphicalNote, note.NoteBeam, openBeams);
                    }
                }
                if (note.NoteTuplet !== undefined && note.PrintObject) {
                    this.handleTuplet(graphicalNote, note.NoteTuplet, openTuplets);
                }
            }

            // handle TabNotes:
            if (graphicalTabVoiceEntry) {
                // notes should be either TabNotes or RestNotes -> add all:
                const graphicalTabNote: GraphicalNote = MusicSheetCalculator.symbolFactory.createNote(
                    note,
                    graphicalTabVoiceEntry,
                    activeClef,
                    octaveShiftValue,
                    this.rules,
                    undefined);
                tabStaffEntry.addGraphicalNoteToListAtCorrectYPosition(graphicalTabVoiceEntry, graphicalTabNote);
                graphicalTabNote.PositionAndShape.calculateBoundingBox();

                if (!this.leadSheet) {
                    if (note.NoteTuplet) {
                        this.handleTuplet(graphicalTabNote, note.NoteTuplet, openTuplets);
                    }
                }
            }
        }
        if (voiceEntry.Articulations.length > 0) {
            this.handleVoiceEntryArticulations(voiceEntry.Articulations, voiceEntry, graphicalStaffEntry);
        }
        if (voiceEntry.TechnicalInstructions.length > 0) {
            this.handleVoiceEntryTechnicalInstructions(voiceEntry.TechnicalInstructions, voiceEntry, graphicalStaffEntry);
        }
        if (voiceEntry.LyricsEntries.size() > 0) {
            this.handleVoiceEntryLyrics(voiceEntry, graphicalStaffEntry, openLyricWords);
        }
        if (voiceEntry.OrnamentContainer) {
            this.handleVoiceEntryOrnaments(voiceEntry.OrnamentContainer, voiceEntry, graphicalStaffEntry);
        }
        return octaveShiftValue;
    }

    protected resetYPositionForLeadSheet(psi: BoundingBox): void {
        if (this.leadSheet) {
            psi.RelativePosition = new PointF2D(psi.RelativePosition.x, 0.0);
        }
    }

    protected layoutVoiceEntries(graphicalStaffEntry: GraphicalStaffEntry, staffIndex: number): void {
        graphicalStaffEntry.PositionAndShape.RelativePosition = new PointF2D(0.0, 0.0);
        if (!this.leadSheet) {
            for (const gve of graphicalStaffEntry.graphicalVoiceEntries) {
                const graphicalNotes: GraphicalNote[] = gve.notes;
                if (graphicalNotes.length === 0) {
                    continue;
                }
                const voiceEntry: VoiceEntry = graphicalNotes[0].sourceNote.ParentVoiceEntry;
                const hasPitchedNote: boolean = graphicalNotes[0].sourceNote.Pitch !== undefined;
                this.layoutVoiceEntry(voiceEntry, graphicalNotes, graphicalStaffEntry, hasPitchedNote);
            }
        }
    }

    protected maxInstrNameLabelLength(): number {
        let maxLabelLength: number = 0.0;
        for (const instrument of this.graphicalMusicSheet.ParentMusicSheet.Instruments) {
            if (instrument.NameLabel?.print && instrument.Voices.length > 0 && instrument.Voices[0].Visible) {
                let renderedLabel: Label = instrument.NameLabel;
                if (!this.rules.RenderPartNames) {
                    renderedLabel = new Label("", renderedLabel.textAlignment, renderedLabel.font);
                }
                const graphicalLabel: GraphicalLabel = new GraphicalLabel(
                    renderedLabel, this.rules.InstrumentLabelTextHeight, TextAlignmentEnum.LeftCenter, this.rules);
                graphicalLabel.setLabelPositionAndShapeBorders();
                maxLabelLength = Math.max(maxLabelLength, graphicalLabel.PositionAndShape.MarginSize.width);
            }
        }
        if (!this.rules.RenderPartNames) {
            return 0;
        }
        return maxLabelLength;
    }

    protected calculateSheetLabelBoundingBoxes(): void {
        const musicSheet: MusicSheet = this.graphicalMusicSheet.ParentMusicSheet;
        const defaultColorTitle: string = this.rules.DefaultColorTitle; // can be undefined => black
        if (musicSheet.Title !== undefined && this.rules.RenderTitle) {
            const title: GraphicalLabel = new GraphicalLabel(musicSheet.Title, this.rules.SheetTitleHeight, TextAlignmentEnum.CenterBottom, this.rules);
            title.Label.IsCreditLabel = true;
            title.Label.colorDefault = defaultColorTitle;
            this.graphicalMusicSheet.Title = title;
            title.setLabelPositionAndShapeBorders();
        } else if (!this.rules.RenderTitle) {
            this.graphicalMusicSheet.Title = undefined; // clear label if rendering it was disabled after last render
        }
        if (musicSheet.Subtitle !== undefined && this.rules.RenderSubtitle) {
            const subtitle: GraphicalLabel = new GraphicalLabel(
                musicSheet.Subtitle, this.rules.SheetSubtitleHeight, TextAlignmentEnum.CenterCenter, this.rules);
            subtitle.Label.IsCreditLabel = true;
            subtitle.Label.colorDefault = defaultColorTitle;
            this.graphicalMusicSheet.Subtitle = subtitle;
            subtitle.setLabelPositionAndShapeBorders();
        } else if (!this.rules.RenderSubtitle) {
            this.graphicalMusicSheet.Subtitle = undefined;
        }
        if (musicSheet.Composer !== undefined && this.rules.RenderComposer) {
            const composer: GraphicalLabel = new GraphicalLabel(
                musicSheet.Composer, this.rules.SheetComposerHeight, TextAlignmentEnum.RightCenter, this.rules);
            composer.Label.IsCreditLabel = true;
            composer.Label.colorDefault = defaultColorTitle;
            this.graphicalMusicSheet.Composer = composer;
            composer.setLabelPositionAndShapeBorders();
        } else if (!this.rules.RenderComposer) {
            this.graphicalMusicSheet.Composer = undefined;
        }
        if (musicSheet.Lyricist !== undefined && this.rules.RenderLyricist) {
            const lyricist: GraphicalLabel = new GraphicalLabel(
                musicSheet.Lyricist, this.rules.SheetAuthorHeight, TextAlignmentEnum.LeftCenter, this.rules);
            lyricist.Label.IsCreditLabel = true;
            lyricist.Label.colorDefault = defaultColorTitle;
            this.graphicalMusicSheet.Lyricist = lyricist;
            lyricist.setLabelPositionAndShapeBorders();
        } else if (!this.rules.RenderLyricist) {
            this.graphicalMusicSheet.Lyricist = undefined;
        }
        if (musicSheet.Copyright !== undefined && this.rules.RenderCopyright) {
            const copyright: GraphicalLabel = new GraphicalLabel(
                musicSheet.Copyright, this.rules.SheetCopyrightHeight, TextAlignmentEnum.CenterBottom, this.rules);
                copyright.Label.IsCreditLabel = true;
                copyright.Label.colorDefault = defaultColorTitle;
            this.graphicalMusicSheet.Copyright = copyright;
            copyright.setLabelPositionAndShapeBorders();
        } else if (!this.rules.RenderCopyright) {
            this.graphicalMusicSheet.Copyright = undefined;
        }
    }

    protected checkMeasuresForWholeRestNotes(): void {
        for (let idx2: number = 0, len2: number = this.musicSystems.length; idx2 < len2; ++idx2) {
            const musicSystem: MusicSystem = this.musicSystems[idx2];
            for (let idx3: number = 0, len3: number = musicSystem.StaffLines.length; idx3 < len3; ++idx3) {
                const staffLine: StaffLine = musicSystem.StaffLines[idx3];
                for (let idx4: number = 0, len4: number = staffLine.Measures.length; idx4 < len4; ++idx4) {
                    const measure: GraphicalMeasure = staffLine.Measures[idx4];
                    if (measure.staffEntries.length === 1) {
                        const gse: GraphicalStaffEntry = measure.staffEntries[0];
                        if (gse.graphicalVoiceEntries.length > 0 && gse.graphicalVoiceEntries[0].notes.length === 1) {
                            const graphicalNote: GraphicalNote = gse.graphicalVoiceEntries[0].notes[0];
                            if (!graphicalNote.sourceNote.Pitch && (new Fraction(1, 2)).lt(graphicalNote.sourceNote.Length)) {
                                this.layoutMeasureWithWholeRest(graphicalNote, gse, measure);
                            }
                        }
                    }
                }
            }
        }
    }

    protected optimizeRestNotePlacement(graphicalStaffEntry: GraphicalStaffEntry, measure: GraphicalMeasure): void {
        if (graphicalStaffEntry.graphicalVoiceEntries.length === 0) {
            return;
        }
        const restNotes: GraphicalNote[] = [];
        const pitchedNotes: GraphicalNote[] = [];
        this.collectRestPlacementNotes(graphicalStaffEntry, restNotes, pitchedNotes);
        if (restNotes.length === 0) {
            return;
        }
        if (pitchedNotes.length > 0 || restNotes.length > 1) {
            this.calculateRestNotesPlacementWithCollisionDetection(graphicalStaffEntry, measure, restNotes, pitchedNotes);
            return;
        }
        const voice1Note1: GraphicalNote = restNotes[0];
        if (voice1Note1 && graphicalStaffEntry !== measure.staffEntries[0] &&
            graphicalStaffEntry !== measure.staffEntries[measure.staffEntries.length - 1]) {
            const staffEntryIndex: number = measure.staffEntries.indexOf(graphicalStaffEntry);
            const previousStaffEntry: GraphicalStaffEntry = measure.staffEntries[staffEntryIndex - 1];
            const nextStaffEntry: GraphicalStaffEntry = measure.staffEntries[staffEntryIndex + 1];
            if (previousStaffEntry.graphicalVoiceEntries.length === 1) {
                const previousNote: GraphicalNote = previousStaffEntry.graphicalVoiceEntries[0].notes[0];
                if (previousNote?.sourceNote?.NoteBeam !== undefined && nextStaffEntry.graphicalVoiceEntries.length === 1) {
                    const nextNote: GraphicalNote = nextStaffEntry.graphicalVoiceEntries[0].notes[0];
                    if (nextNote?.sourceNote?.NoteBeam !== undefined &&
                        previousNote.sourceNote.NoteBeam === nextNote.sourceNote.NoteBeam) {
                        this.calculateRestNotePlacementWithinGraphicalBeam(
                            graphicalStaffEntry, voice1Note1, previousNote,
                            nextStaffEntry, nextNote
                        );
                        graphicalStaffEntry.PositionAndShape.calculateBoundingBox();
                    }
                }
            }
        }
    }

    protected getRelativePositionInStaffLineFromTimestamp(
        timestamp: Fraction, verticalIndex: number, staffLine: StaffLine,
        multiStaffInstrument: boolean, firstVisibleMeasureRelativeX: number = 0.0,
        useLeftStaffEntryBorder: boolean = false
    ): PointF2D {
        let relative: PointF2D = new PointF2D();
        let leftStaffEntry: GraphicalStaffEntry = undefined;
        let rightStaffEntry: GraphicalStaffEntry = undefined;
        const numEntries: number = this.graphicalMusicSheet.VerticalGraphicalStaffEntryContainers.length;
        const index: number = this.graphicalMusicSheet.GetInterpolatedIndexInVerticalContainers(timestamp);
        const leftIndex: number = Math.min(Math.floor(index), numEntries - 1);
        const rightIndex: number = Math.min(Math.ceil(index), numEntries - 1);
        if (leftIndex < 0 || verticalIndex < 0) {
            return relative;
        }
        leftStaffEntry = this.getFirstLeftNotNullStaffEntryFromContainer(leftIndex, verticalIndex, multiStaffInstrument);
        rightStaffEntry = this.getFirstRightNotNullStaffEntryFromContainer(rightIndex, verticalIndex, multiStaffInstrument);
        if (leftStaffEntry && rightStaffEntry) {
            let measureRelativeX: number = leftStaffEntry.parentMeasure.PositionAndShape.RelativePosition.x;
            if (firstVisibleMeasureRelativeX > 0) {
                measureRelativeX = firstVisibleMeasureRelativeX;
            }
            let leftX: number = leftStaffEntry.PositionAndShape.RelativePosition.x + measureRelativeX;
            let rightX: number = rightStaffEntry.PositionAndShape.RelativePosition.x + rightStaffEntry.parentMeasure.PositionAndShape.RelativePosition.x;
            const endAfterRightStaffEntry: boolean = timestamp.RealValue > rightStaffEntry.getAbsoluteTimestamp().RealValue;
            // endAfterRightStaffEntry is an unfortunate case where the timestamp isn't correct for the last note in the piece,
            //   see test_wedge_diminuendo_duplicated.musicxml
            if (firstVisibleMeasureRelativeX > 0) {
                rightX = rightStaffEntry.PositionAndShape.RelativePosition.x + measureRelativeX;
            } else if (useLeftStaffEntryBorder &&
                (leftStaffEntry.getAbsoluteTimestamp().RealValue === timestamp.RealValue || endAfterRightStaffEntry)
            ) {
                leftX = leftStaffEntry.PositionAndShape.RelativePosition.x + leftStaffEntry.PositionAndShape.BorderLeft + measureRelativeX;
                rightX = leftX;
            }
            let timestampQuotient: number = 0.0;
            if (leftStaffEntry !== rightStaffEntry) {
                const leftTimestamp: Fraction = leftStaffEntry.getAbsoluteTimestamp();
                const rightTimestamp: Fraction = rightStaffEntry.getAbsoluteTimestamp();
                const leftDifference: Fraction = Fraction.minus(timestamp, leftTimestamp);
                timestampQuotient = leftDifference.RealValue / Fraction.minus(rightTimestamp, leftTimestamp).RealValue;
            }
            if (leftStaffEntry.parentMeasure.ParentStaffLine !== rightStaffEntry.parentMeasure.ParentStaffLine) {
                if (leftStaffEntry.parentMeasure.ParentStaffLine === staffLine) {
                    rightX = staffLine.PositionAndShape.Size.width;
                } else {
                    leftX = staffLine.PositionAndShape.RelativePosition.x;
                }
            }
            relative = new PointF2D(leftX + (rightX - leftX) * timestampQuotient, 0.0);
        }
        return relative;
    }

    protected getRelativeXPositionFromTimestamp(timestamp: Fraction): number {
        const numEntries: number = this.graphicalMusicSheet.VerticalGraphicalStaffEntryContainers.length;
        const index: number = this.graphicalMusicSheet.GetInterpolatedIndexInVerticalContainers(timestamp);
        const discreteIndex: number = Math.max(0, Math.min(Math.round(index), numEntries - 1));
        const gse: GraphicalStaffEntry = this.graphicalMusicSheet.VerticalGraphicalStaffEntryContainers[discreteIndex].getFirstNonNullStaffEntry();
        const posX: number = gse.PositionAndShape.RelativePosition.x + gse.parentMeasure.PositionAndShape.RelativePosition.x;
        return posX;
    }

    protected calculatePageLabels(page: GraphicalMusicPage): void {
        // The PositionAndShape child elements of page need to be manually connected to the lyricist, composer, subtitle, etc.
        // because the page is only available now

        // fix width of SVG, sheet and horizontal scroll bar being too long (~32767 = SheetMaximumWidth) for single line scores
        if (this.rules.RenderSingleHorizontalStaffline) {
            //page.PositionAndShape.BorderRight = page.PositionAndShape.Size.width + this.rules.PageRightMargin;
            page.PositionAndShape.calculateBoundingBox([GraphicalMeasure.name]); // ignore measures, whose bounding boxes somehow get messed up otherwise
            // note: "GraphicalMeasure" instead of GraphicalMeasure.name doesn't work with minified builds (they change class names)
            // note: calculateBoundingBox by default changes measure.PositionAndShape.Size.width for some reason,
            //   inaccurate for RenderSingleHorizontalStaffline, e.g. the cursor type 3 that highlights the whole measure will get wrong width
            //   correct width was set previously via MusicSystemBuilder.setMeasureWidth().

            // limit SVG and scroll bar width so it's not ~32767 (SheetMaximumWidth):
            this.graphicalMusicSheet.ParentMusicSheet.pageWidth = page.PositionAndShape.Size.width;
            // page.PositionAndShape.BorderRight = page.PositionAndShape.Size.width; // doesn't seem to affect anything
        }

        let firstSystemAbsoluteTopMargin: number = 10;
        let lastSystemAbsoluteBottomMargin: number = -1;
        if (page.MusicSystems.length > 0) {
            const firstMusicSystem: MusicSystem = page.MusicSystems[0];
            firstSystemAbsoluteTopMargin = firstMusicSystem.PositionAndShape.RelativePosition.y + firstMusicSystem.PositionAndShape.BorderTop;
            const lastMusicSystem: MusicSystem = page.MusicSystems[page.MusicSystems.length - 1];
            lastSystemAbsoluteBottomMargin = lastMusicSystem.PositionAndShape.RelativePosition.y + lastMusicSystem.PositionAndShape.BorderBottom;
        }
        //const firstStaffLine: StaffLine = this.graphicalMusicSheet.MusicPages[0].MusicSystems[0].StaffLines[0];
        const title: GraphicalLabel = this.graphicalMusicSheet.Title;
        if (title && this.rules.RenderTitle) {
            title.PositionAndShape.Parent = page.PositionAndShape;
            //title.PositionAndShape.Parent = firstStaffLine.PositionAndShape;
            const relative: PointF2D = new PointF2D();
            relative.x = this.graphicalMusicSheet.ParentMusicSheet.pageWidth / 2;
            if (this.rules.RenderSingleHorizontalStaffline) {
                relative.x = Math.max(relative.x, title.PositionAndShape.Size.width);
            }
            //relative.x = firstStaffLine.PositionAndShape.RelativePosition.x + firstStaffLine.PositionAndShape.Size.width / 2; // half of first staffline width
            relative.y = this.rules.TitleTopDistance + this.rules.SheetTitleHeight;
            title.PositionAndShape.RelativePosition = relative;
            page.Labels.push(title);
        }
        if (this.graphicalMusicSheet.Subtitle && this.rules.RenderTitle && this.rules.RenderSubtitle) {
            const subtitle: GraphicalLabel = this.graphicalMusicSheet.Subtitle;
            // subtitle.PositionAndShape.Parent = firstStaffLine.PositionAndShape;
            subtitle.PositionAndShape.Parent = page.PositionAndShape;
            const relative: PointF2D = new PointF2D();
            relative.x = this.graphicalMusicSheet.ParentMusicSheet.pageWidth / 2;
            if (this.rules.RenderSingleHorizontalStaffline) {
                relative.x = title.PositionAndShape.RelativePosition.x; //Math.max(relative.x, title.PositionAndShape.Size.width);
            }
            //relative.x = firstStaffLine.PositionAndShape.RelativePosition.x + firstStaffLine.PositionAndShape.Size.width / 2; // half of first staffline width
            relative.y = this.rules.TitleTopDistance + this.rules.SheetTitleHeight + this.rules.SheetMinimumDistanceBetweenTitleAndSubtitle;
            const lines: number = subtitle.TextLines?.length;
            if (lines > 1) { // Don't want to affect existing behavior. but this doesn't check bboxes for clip
                relative.y += subtitle.PositionAndShape.BorderBottom * (lines - 1) / (lines);
            }
            subtitle.PositionAndShape.RelativePosition = relative;
            page.Labels.push(subtitle);
        }
        // Get the first system, first staffline skybottomcalculator
        // const topStaffline: StaffLine = page.MusicSystems[0].StaffLines[0];
        // const skyBottomLineCalculator: SkyBottomLineCalculator = topStaffline.SkyBottomLineCalculator;
        //   we don't need a skybottomcalculator currently, labels are put above system skyline anyways.
        const composer: GraphicalLabel = this.graphicalMusicSheet.Composer;
        let composerRelativeY: number;
        if (composer && this.rules.RenderComposer) {
            composer.PositionAndShape.Parent = page.PositionAndShape; // if using pageWidth. (which can currently be too wide) TODO fix pageWidth (#578)
            //composer.PositionAndShape.Parent = topStaffline.PositionAndShape; // if using firstStaffLine...width.
            //      y-collision problems, harder to y-align with lyrics
            composer.setLabelPositionAndShapeBorders();
            const relative: PointF2D = new PointF2D();
            //const firstStaffLineEndX: number = this.rules.PageLeftMargin + this.rules.SystemLeftMargin + this.rules.left
            //    firstStaffLine.PositionAndShape.RelativePosition.x + firstStaffLine.PositionAndShape.Size.width;
            //relative.x = Math.min(this.graphicalMusicSheet.ParentMusicSheet.pageWidth - this.rules.PageRightMargin,
            //  firstStaffLineEndX); // awkward with 2-bar score
            relative.x = this.graphicalMusicSheet.ParentMusicSheet.pageWidth - this.rules.PageRightMargin;
            // if (this.rules.RenderSingleHorizontalStaffline) {
            //     relative.x = page.PositionAndShape.BorderMarginLeft + title.PositionAndShape.Size.width * 2;
            // }
            //relative.x = firstStaffLine.PositionAndShape.Size.width;
            //when this is less, goes higher.
            //So 0 is top of the sheet, 22 or so is touching the music system margin

            relative.y = firstSystemAbsoluteTopMargin;
            //relative.y = - this.rules.SystemComposerDistance;
            //relative.y = -firstStaffLine.PositionAndShape.Size.height;
            // TODO only add measure label height if rendering labels and composer measure has label
            // TODO y-align with lyricist? which is harder if they have different bbox parents (page and firstStaffLine).
            // when the pageWidth gets fixed, we could use page as parent again.

            //Sufficient for now to just use the longest composer entry instead of bottom.
            //Otherwise we need to construct a 'bottom line' for the text block
            // const endX: number = topStaffline.PositionAndShape.BorderMarginRight;
            // const startX: number = endX - composer.PositionAndShape.Size.width;
            // const currentMin: number = skyBottomLineCalculator.getSkyLineMinInRange(startX, endX);

            relative.y -= this.rules.SystemComposerDistance;
            const lines: number = composer.TextLines?.length;
            if (lines > 1) { //Don't want to affect existing behavior. but this doesn't check bboxes for clip
                relative.y -= composer.PositionAndShape.BorderBottom * (lines - 1) / (lines);
            }
            //const newSkylineY: number = currentMin; // don't add composer label height to skyline
            //- firstSystemAbsoluteTopMargin - this.rules.SystemComposerDistance - composer.PositionAndShape.MarginSize.height;
            //skyBottomLineCalculator.updateSkyLineInRange(startX, endX, newSkylineY); // this can fix skyline for generateImages for some reason
            composerRelativeY = relative.y; // for lyricist label

            composer.PositionAndShape.RelativePosition = relative;
            page.Labels.push(composer);
        }
        const lyricist: GraphicalLabel = this.graphicalMusicSheet.Lyricist;
        if (lyricist && this.rules.RenderLyricist) {
            lyricist.PositionAndShape.Parent = page.PositionAndShape;
            lyricist.setLabelPositionAndShapeBorders();
            const relative: PointF2D = new PointF2D();
            relative.x = this.rules.PageLeftMargin;
            relative.y = firstSystemAbsoluteTopMargin;
            // const startX: number = topStaffline.PositionAndShape.BorderMarginLeft - relative.x;
            // const endX: number = startX + lyricist.PositionAndShape.Size.width;
            // const currentMin: number = skyBottomLineCalculator.getSkyLineMinInRange(startX, endX);

            relative.y -= this.rules.SystemLyricistDistance;
            relative.y += lyricist.PositionAndShape.BorderBottom;
            relative.y = Math.min(relative.y, composerRelativeY ?? Number.MAX_SAFE_INTEGER);
            // same height as composer label (at least not lower). ?? prevents undefined -> Math.min returns NaN

            //skyBottomLineCalculator.updateSkyLineInRange(startX, endX, currentMin - lyricist.PositionAndShape.MarginSize.height);
            //relative.y = Math.max(relative.y, composer.PositionAndShape.RelativePosition.y);
            lyricist.PositionAndShape.RelativePosition = relative;
            page.Labels.push(lyricist);
        }
        const copyright: GraphicalLabel = this.graphicalMusicSheet.Copyright;
        if (copyright && this.rules.RenderCopyright) {
            copyright.PositionAndShape.Parent = page.PositionAndShape;
            copyright.setLabelPositionAndShapeBorders();
            const relative: PointF2D = new PointF2D();
            relative.x = page.PositionAndShape.Size.width / 2;
            relative.y = lastSystemAbsoluteBottomMargin + this.rules.SheetCopyrightMargin;
            relative.y -= copyright.PositionAndShape.BorderTop;
            copyright.PositionAndShape.RelativePosition = relative;
            page.Labels.push(copyright);
        }
        // we need to do this again to not cut off the title for short scores:
        //   (and fix SVG and horizontal scroll bar width)
        if (this.rules.RenderSingleHorizontalStaffline) {
            //page.PositionAndShape.BorderRight = page.PositionAndShape.Size.width + this.rules.PageRightMargin;
            page.PositionAndShape.calculateBoundingBox([GraphicalMeasure.name]); // ignore measures, whose bounding boxes somehow get messed up otherwise
            // note: calculateBoundingBox by default changes measure.PositionAndShape.Size.width for some reason,
            //   inaccurate for RenderSingleHorizontalStaffline, e.g. the cursor type 3 that highlights the whole measure will get wrong width
            //   correct width was set previously via MusicSystemBuilder.setMeasureWidth().

            // limit SVG and scroll bar width so it's not ~32767 (SheetMaximumWidth):
            this.graphicalMusicSheet.ParentMusicSheet.pageWidth = page.PositionAndShape.Size.width;
            // page.PositionAndShape.BorderRight = page.PositionAndShape.Size.width; // doesn't seem to affect anything
        }
    }

    protected createGraphicalTies(): void {
        for (let measureIndex: number = 0; measureIndex < this.graphicalMusicSheet.ParentMusicSheet.SourceMeasures.length; measureIndex++) {
            const sourceMeasure: SourceMeasure = this.graphicalMusicSheet.ParentMusicSheet.SourceMeasures[measureIndex];
            for (let staffIndex: number = 0; staffIndex < sourceMeasure.CompleteNumberOfStaves; staffIndex++) {
                for (let j: number = 0; j < sourceMeasure.VerticalSourceStaffEntryContainers.length; j++) {
                    const sourceStaffEntry: SourceStaffEntry = sourceMeasure.VerticalSourceStaffEntryContainers[j].StaffEntries[staffIndex];
                    if (sourceStaffEntry) {
                        const startStaffEntry: GraphicalStaffEntry = this.graphicalMusicSheet.findGraphicalStaffEntryFromMeasureList(
                            staffIndex, measureIndex, sourceStaffEntry
                        );
                        if (startStaffEntry) {
                            startStaffEntry.GraphicalTies.clear(); // don't duplicate ties when calling render() again
                            startStaffEntry.ties.clear();
                        }

                        for (let idx: number = 0, len: number = sourceStaffEntry.VoiceEntries.length; idx < len; ++idx) {
                            const voiceEntry: VoiceEntry = sourceStaffEntry.VoiceEntries[idx];
                            for (let idx2: number = 0, len2: number = voiceEntry.Notes.length; idx2 < len2; ++idx2) {
                                const note: Note = voiceEntry.Notes[idx2];
                                if (note.NoteTie) {
                                    const tie: Tie = note.NoteTie;
                                    if (note === note.NoteTie.Notes.last()) {
                                        continue; // nothing to do on last note. don't create last tie twice.
                                    }
                                    if (startStaffEntry) {
                                        for (const gTie of startStaffEntry.GraphicalTies) {
                                            if (gTie.Tie === tie) {
                                                continue; // don't handle the same tie on the same startStaffEntry twice
                                            }
                                        }
                                    }
                                    this.handleTie(tie, startStaffEntry, staffIndex, measureIndex);
                                }
                            }
                        }
                        this.setTieDirections(startStaffEntry);
                    }
                }
            }
        }
    }

    private handleTie(tie: Tie, startGraphicalStaffEntry: GraphicalStaffEntry, staffIndex: number, measureIndex: number): void {
        if (!startGraphicalStaffEntry) {
            // console.log('tie not found in measure number ' + measureIndex - 1);
            return;
        }
        startGraphicalStaffEntry.ties.push(tie);

        let startGse: GraphicalStaffEntry = startGraphicalStaffEntry;
        let startNote: GraphicalNote = undefined;
        let endGse: GraphicalStaffEntry = undefined;
        let endNote: GraphicalNote = undefined;
        for (let i: number = 1; i < tie.Notes.length; i++) {
            startNote = startGse.findTieGraphicalNoteFromNote(tie.Notes[i - 1]);
            endGse = this.graphicalMusicSheet.GetGraphicalFromSourceStaffEntry(tie.Notes[i].ParentStaffEntry);
            if (!endGse) {
                continue;
            }
            endNote = endGse.findTieGraphicalNoteFromNote(tie.Notes[i]);
            if (startNote !== undefined && endNote !== undefined && endGse) {
                if (!startNote.sourceNote.PrintObject || !endNote.sourceNote.PrintObject) {
                    continue;
                }
                const graphicalTie: GraphicalTie = this.createGraphicalTie(tie, startGse, endGse, startNote, endNote);
                startGse.GraphicalTies.push(graphicalTie);
                if (this.staffEntriesWithGraphicalTies.indexOf(startGse) >= 0) {
                    this.staffEntriesWithGraphicalTies.push(startGse);
                }
            }
            startGse = endGse;
        }
    }

    private setTieDirections(staffEntry: GraphicalStaffEntry): void {
        if (!staffEntry) {
            return;
        }
        const ties: Tie[] = staffEntry.ties;
        if (ties.length === 1) {
            const tie: Tie = ties[0];
            if (tie.TieDirection === PlacementEnum.NotYetDefined) {
                const voiceId: number = tie.Notes[0].ParentVoiceEntry.ParentVoice.VoiceId;
                // put ties of second voices (e.g. 2 for right hand, 6 left hand) below by default
                //   TODO could be more precise but also more complex by checking lower notes, other notes, etc.
                if (voiceId === 2 || voiceId === 6) {
                    tie.TieDirection = PlacementEnum.Below;
                }
            }
        }
        if (ties.length > 1) {
            let highestNote: Note = undefined;
            for (const gseTie of ties) {
                const tieNote: Note = gseTie.Notes[0];
                if (!highestNote || tieNote.Pitch.getHalfTone() > highestNote.Pitch.getHalfTone()) {
                    highestNote = tieNote;
                }
            }
            for (const gseTie of ties) {
                if (gseTie.TieDirection === PlacementEnum.NotYetDefined) { // only set/change if not already set by xml
                    if (gseTie.Notes[0] === highestNote) {
                        gseTie.TieDirection = PlacementEnum.Above;
                    } else {
                        gseTie.TieDirection = PlacementEnum.Below;
                    }
                }
            }
        }
    }

    private createAccidentalCalculators(): AccidentalCalculator[] {
        const accidentalCalculators: AccidentalCalculator[] = [];
        const firstSourceMeasure: SourceMeasure = this.graphicalMusicSheet.ParentMusicSheet.getFirstSourceMeasure();
        if (firstSourceMeasure) {
            const transposeHalftones: number = this.graphicalMusicSheet.ParentMusicSheet.Transpose;
            for (let i: number = 0; i < firstSourceMeasure.CompleteNumberOfStaves; i++) {
                const accidentalCalculator: AccidentalCalculator = new AccidentalCalculator();
                accidentalCalculators.push(accidentalCalculator);
                accidentalCalculator.Transpose = transposeHalftones;
                if (firstSourceMeasure.FirstInstructionsStaffEntries[i]) {
                    for (let idx: number = 0, len: number = firstSourceMeasure.FirstInstructionsStaffEntries[i].Instructions.length; idx < len; ++idx) {
                        const abstractNotationInstruction: AbstractNotationInstruction = firstSourceMeasure.FirstInstructionsStaffEntries[i].Instructions[idx];
                        if (abstractNotationInstruction instanceof KeyInstruction) {
                            // Create a new KeyInstruction using keyTypeOriginal to ensure correct starting point (#1383)
                            // This ensures that when transpose=0, we get the original key (e.g., C major)
                            // rather than a previously transposed key (e.g., Db major from transpose=1)
                            const originalKey: KeyInstruction = <KeyInstruction>abstractNotationInstruction;
                            const key: KeyInstruction = new KeyInstruction(originalKey.Parent, originalKey.keyTypeOriginal, originalKey.Mode);
                            // Then transpose if needed (skip percussion instruments)
                            const staff: Staff = this.graphicalMusicSheet.ParentMusicSheet.Staves[i];
                            if (transposeHalftones !== 0 &&
                                staff?.ParentInstrument?.MidiInstrumentId !== MidiInstrument.Percussion &&
                                MusicSheetCalculator.transposeCalculator) {
                                MusicSheetCalculator.transposeCalculator.transposeKey(key, transposeHalftones);
                            }
                            accidentalCalculator.ActiveKeyInstruction = key;
                        }
                    }
                }
            }
        }
        return accidentalCalculators;
    }

    private calculateVerticalContainersList(): void {
        const numberOfEntries: number = this.graphicalMusicSheet.MeasureList[0].length;
        for (let i: number = 0; i < this.graphicalMusicSheet.MeasureList.length; i++) {
            for (let j: number = 0; j < numberOfEntries; j++) {
                const measure: GraphicalMeasure = this.graphicalMusicSheet.MeasureList[i][j];
                if (!measure) {
                    continue;
                }
                for (let idx: number = 0, len: number = measure.staffEntries.length; idx < len; ++idx) {
                    const graphicalStaffEntry: GraphicalStaffEntry = measure.staffEntries[idx];
                    const verticalContainer: VerticalGraphicalStaffEntryContainer =
                        this.graphicalMusicSheet.getOrCreateVerticalContainer(graphicalStaffEntry.getAbsoluteTimestamp());
                    if (verticalContainer) {
                        verticalContainer.StaffEntries[j] = graphicalStaffEntry;
                        graphicalStaffEntry.parentVerticalContainer = verticalContainer;
                    }
                }
            }
        }
    }

    private setIndicesToVerticalGraphicalContainers(): void {
        for (let i: number = 0; i < this.graphicalMusicSheet.VerticalGraphicalStaffEntryContainers.length; i++) {
            this.graphicalMusicSheet.VerticalGraphicalStaffEntryContainers[i].Index = i;
        }
    }

    private createGraphicalMeasuresForSourceMeasure(sourceMeasure: SourceMeasure, accidentalCalculators: AccidentalCalculator[],
                                                    openLyricWords: LyricWord[],
                                                    openOctaveShifts: OctaveShiftParams[], activeClefs: ClefInstruction[]): GraphicalMeasure[] {
        this.initGraphicalMeasuresCreation();
        const verticalMeasureList: GraphicalMeasure[] = []; // (VexFlowMeasure, extends GraphicalMeasure)
        const openBeams: Beam[] = [];
        const openTuplets: Tuplet[] = [];
        const staffEntryLinks: StaffEntryLink[] = [];
        let restInAllGraphicalMeasures: boolean = true;
        for (let staffIndex: number = 0; staffIndex < sourceMeasure.CompleteNumberOfStaves; staffIndex++) {
            const measure: GraphicalMeasure = this.createGraphicalMeasure( // (VexFlowMeasure)
                sourceMeasure, openTuplets, openBeams,
                accidentalCalculators[staffIndex], activeClefs, openOctaveShifts, openLyricWords, staffIndex, staffEntryLinks
            );
            restInAllGraphicalMeasures = restInAllGraphicalMeasures && measure.hasOnlyRests;
            verticalMeasureList.push(measure);
        }
        sourceMeasure.allRests = restInAllGraphicalMeasures;
        sourceMeasure.VerticalMeasureList = verticalMeasureList; // much easier way to link sourceMeasure to graphicalMeasures than Dictionary
        //this.graphicalMusicSheet.sourceToGraphicalMeasureLinks.setValue(sourceMeasure, verticalMeasureList); // overwrites entries because:
        //this.graphicalMusicSheet.sourceToGraphicalMeasureLinks[sourceMeasure] = verticalMeasureList; // can't use SourceMeasure as key.
        // to save the reference by dictionary we would need two Dictionaries, id -> sourceMeasure and id -> GraphicalMeasure.
        return verticalMeasureList;
    }

    private createGraphicalMeasure(sourceMeasure: SourceMeasure, openTuplets: Tuplet[], openBeams: Beam[],
                                   accidentalCalculator: AccidentalCalculator, activeClefs: ClefInstruction[],
                                   openOctaveShifts: OctaveShiftParams[], openLyricWords: LyricWord[], staffIndex: number,
                                   staffEntryLinks: StaffEntryLink[]): GraphicalMeasure {
        const staff: Staff = this.graphicalMusicSheet.ParentMusicSheet.getStaffFromIndex(staffIndex);
        let measure: GraphicalMeasure = undefined;
        if (activeClefs[staffIndex].ClefType === ClefEnum.TAB || staff.isTab) {
            staff.isTab = true;
            measure = MusicSheetCalculator.symbolFactory.createTabStaffMeasure(sourceMeasure, staff);
        } else if (sourceMeasure.multipleRestMeasures && this.rules.RenderMultipleRestMeasures) {
            measure = MusicSheetCalculator.symbolFactory.createMultiRestMeasure(sourceMeasure, staff);
        } else if (sourceMeasure.multipleRestMeasureNumber > 1) {
            return undefined; // don't need to create a graphical measure that is within a multiple rest measure
        } else {
            measure = MusicSheetCalculator.symbolFactory.createGraphicalMeasure(sourceMeasure, staff);
        }
        measure.hasError = sourceMeasure.getErrorInMeasure(staffIndex);
        // check for key instruction changes
        if (sourceMeasure.FirstInstructionsStaffEntries[staffIndex]) {
            for (let idx: number = 0, len: number = sourceMeasure.FirstInstructionsStaffEntries[staffIndex].Instructions.length; idx < len; ++idx) {
                const instruction: AbstractNotationInstruction = sourceMeasure.FirstInstructionsStaffEntries[staffIndex].Instructions[idx];
                if (instruction instanceof KeyInstruction) {
                    // Create a new KeyInstruction using keyTypeOriginal to ensure correct starting point (#1383)
                    // This ensures that when transpose=0, we get the original key (e.g., C major)
                    // rather than a previously transposed key (e.g., Db major from transpose=1)
                    const key: KeyInstruction = new KeyInstruction(instruction.Parent, instruction.keyTypeOriginal, instruction.Mode);
                    const transposeHalftones: number = measure.getTransposedHalftones();
                    if (transposeHalftones !== 0 && MusicSheetCalculator.transposeCalculator === undefined) {
                        log.info("[OSMD] transpose requested, but TransposeCalculator undefined. Use osmd.TransposeCalculator = new TransposeCalculator()");
                    }
                    if (transposeHalftones !== 0 &&
                        measure.ParentStaff.ParentInstrument.MidiInstrumentId !== MidiInstrument.Percussion &&
                        MusicSheetCalculator.transposeCalculator) {
                        MusicSheetCalculator.transposeCalculator.transposeKey(
                            key, transposeHalftones
                        );
                    }
                    accidentalCalculator.ActiveKeyInstruction = key;
                }
            }
        }
        // check for octave shifts
        const octaveShifts: MultiExpression[] = [];
        for (let idx: number = 0, len: number = sourceMeasure.StaffLinkedExpressions[staffIndex].length; idx < len; ++idx) {
            const multiExpression: MultiExpression = sourceMeasure.StaffLinkedExpressions[staffIndex][idx];
            let targetOctaveShift: OctaveShift;
            if (multiExpression.OctaveShiftStart) {
                targetOctaveShift = multiExpression.OctaveShiftStart;
            } else if (multiExpression.OctaveShiftEnd) {
                // also check for octave shift that is ending here but starting in earlier measure, see test_octaveshift_notes_shifted_octave_shift_end.musicxml
                targetOctaveShift = multiExpression.OctaveShiftEnd;
            }
            if (targetOctaveShift) {
                octaveShifts.push(multiExpression);
                const openOctaveShift: OctaveShift = targetOctaveShift;
                let absoluteEnd: Fraction = openOctaveShift?.ParentEndMultiExpression?.AbsoluteTimestamp;
                if (!openOctaveShift?.ParentEndMultiExpression) {
                    const measureEndTimestamp: Fraction = Fraction.plus(sourceMeasure.AbsoluteTimestamp, sourceMeasure.Duration);
                    absoluteEnd = measureEndTimestamp;
                    // TODO better handling if end expression missing
                    // old comment:
                    // TODO check if octaveshift end exists, otherwise set to last measure end. only necessary if xml was cut manually and is incomplete
                }
                openOctaveShifts[staffIndex] = new OctaveShiftParams(
                    openOctaveShift, openOctaveShift.ParentStartMultiExpression.AbsoluteTimestamp,
                    //openOctaveShift, multiExpression?.AbsoluteTimestamp,
                    absoluteEnd
                );
            }
        }
        // create GraphicalStaffEntries - always check for possible null Entry
        for (let entryIndex: number = 0; entryIndex < sourceMeasure.VerticalSourceStaffEntryContainers.length; entryIndex++) {
            const sourceStaffEntry: SourceStaffEntry = sourceMeasure.VerticalSourceStaffEntryContainers[entryIndex].StaffEntries[staffIndex];
            // is there a SourceStaffEntry at this Index
            if (sourceStaffEntry) {
                // a SourceStaffEntry exists
                // is there an inStaff ClefInstruction? -> update activeClef
                for (let idx: number = 0, len: number = sourceStaffEntry.Instructions.length; idx < len; ++idx) {
                    const abstractNotationInstruction: AbstractNotationInstruction = sourceStaffEntry.Instructions[idx];
                    if (abstractNotationInstruction instanceof ClefInstruction && activeClefs[staffIndex]?.ClefType !== ClefEnum.TAB) {
                        // if activeClef is TAB, changing it can make the current/next tab measure look like a classical measure. See #1592
                        activeClefs[staffIndex] = <ClefInstruction>abstractNotationInstruction;
                    }
                }
                // create new GraphicalStaffEntry
                const graphicalStaffEntry: GraphicalStaffEntry = MusicSheetCalculator.symbolFactory.createStaffEntry(sourceStaffEntry, measure);
                if (entryIndex < measure.staffEntries.length) {
                    // a GraphicalStaffEntry has been inserted already at this Index (from Tie)
                    measure.addGraphicalStaffEntryAtTimestamp(graphicalStaffEntry);
                } else {
                    measure.addGraphicalStaffEntry(graphicalStaffEntry);
                }

                const linkedNotes: Note[] = [];
                if (sourceStaffEntry.Link) {
                    sourceStaffEntry.findLinkedNotes(linkedNotes);
                    this.handleStaffEntryLink(graphicalStaffEntry, staffEntryLinks);
                }
                // check for possible OctaveShift
                let octaveShiftValue: OctaveEnum = OctaveEnum.NONE;
                let activeOctaveShift: OctaveShift;
                if (openOctaveShifts[staffIndex]) {
                    if (openOctaveShifts[staffIndex].getAbsoluteStartTimestamp.lte(sourceStaffEntry.AbsoluteTimestamp) &&
                        sourceStaffEntry.AbsoluteTimestamp.lte(openOctaveShifts[staffIndex].getAbsoluteEndTimestamp)) {
                        octaveShiftValue = openOctaveShifts[staffIndex].getOpenOctaveShift.Type;
                        activeOctaveShift = openOctaveShifts[staffIndex].getOpenOctaveShift;
                    }
                }
                if (octaveShiftValue === OctaveEnum.NONE) {
                    // check for existing octave shifts outside openOctaveShifts
                    for (const octaveShift of octaveShifts) {
                        let targetOctaveShift: OctaveShift;
                        if (octaveShift.OctaveShiftStart) {
                            targetOctaveShift = octaveShift.OctaveShiftStart;
                        } else if (octaveShift.OctaveShiftEnd) {
                            targetOctaveShift = octaveShift.OctaveShiftEnd;
                        }
                        if (targetOctaveShift?.ParentStartMultiExpression?.AbsoluteTimestamp.lte(sourceStaffEntry.AbsoluteTimestamp) &&
                            !targetOctaveShift.ParentEndMultiExpression?.AbsoluteTimestamp.lt(sourceStaffEntry.AbsoluteTimestamp)) {
                                octaveShiftValue = targetOctaveShift.Type;
                                activeOctaveShift = targetOctaveShift;
                                break;
                            }
                    }
                }
                // for each visible Voice create the corresponding GraphicalNotes
                for (let idx: number = 0, len: number = sourceStaffEntry.VoiceEntries.length; idx < len; ++idx) {
                    const voiceEntry: VoiceEntry = sourceStaffEntry.VoiceEntries[idx];
                    // When an octave shift stop falls between grace notes at the same timestamp,
                    // turn off the shift for VoiceEntries parsed after the stop.
                    if (octaveShiftValue !== OctaveEnum.NONE && activeOctaveShift &&
                        activeOctaveShift.endVoiceEntryIndex > 0 && // without this, last note in bracket is drawn an octave too high
                        activeOctaveShift.ParentEndMultiExpression?.AbsoluteTimestamp.Equals(sourceStaffEntry.AbsoluteTimestamp) &&
                        idx >= activeOctaveShift.endVoiceEntryIndex) {
                        octaveShiftValue = OctaveEnum.NONE;
                        activeOctaveShift = undefined;
                    }
                    // Normal Notes...
                    octaveShiftValue = this.handleVoiceEntry(
                        voiceEntry, graphicalStaffEntry,
                        accidentalCalculator, openLyricWords,
                        activeClefs[staffIndex], openTuplets,
                        openBeams, octaveShiftValue, staffIndex,
                        linkedNotes, sourceStaffEntry
                    );
                }
                // SourceStaffEntry has inStaff ClefInstruction -> create graphical clef
                for (const instruction of sourceStaffEntry.Instructions) {
                    if (instruction instanceof ClefInstruction) {
                        MusicSheetCalculator.symbolFactory.createInStaffClef(
                            graphicalStaffEntry, instruction as ClefInstruction
                        );
                        break;
                    }
                }
                if (this.rules.RenderChordSymbols && sourceStaffEntry.ChordContainers?.length > 0) {
                    sourceStaffEntry.ParentStaff.ParentInstrument.HasChordSymbols = true;
                    MusicSheetCalculator.symbolFactory.createChordSymbols(
                        sourceStaffEntry,
                        graphicalStaffEntry,
                        accidentalCalculator.ActiveKeyInstruction,
                        this.graphicalMusicSheet.ParentMusicSheet.Transpose);
                }
            }
        }

        accidentalCalculator.doCalculationsAtEndOfMeasure();
        // update activeClef given at end of measure if needed
        if (sourceMeasure.LastInstructionsStaffEntries[staffIndex]) {
            const lastStaffEntry: SourceStaffEntry = sourceMeasure.LastInstructionsStaffEntries[staffIndex];
            for (let idx: number = 0, len: number = lastStaffEntry.Instructions.length; idx < len; ++idx) {
                const abstractNotationInstruction: AbstractNotationInstruction = lastStaffEntry.Instructions[idx];
                if (abstractNotationInstruction instanceof ClefInstruction && activeClefs[staffIndex]?.ClefType !== ClefEnum.TAB) {
                    activeClefs[staffIndex] = <ClefInstruction>abstractNotationInstruction;
                }
            }
        }
        for (let idx: number = 0, len: number = sourceMeasure.StaffLinkedExpressions[staffIndex].length; idx < len; ++idx) {
            const multiExpression: MultiExpression = sourceMeasure.StaffLinkedExpressions[staffIndex][idx];
            if (multiExpression.OctaveShiftEnd !== undefined && openOctaveShifts[staffIndex] !== undefined &&
                multiExpression.OctaveShiftEnd === openOctaveShifts[staffIndex].getOpenOctaveShift) {
                    openOctaveShifts[staffIndex] = undefined;
            }
        }
        // check wantedStemDirections of beam notes at end of measure (e.g. for beam with grace notes)
        for (const staffEntry of measure.staffEntries) {
            for (const voiceEntry of staffEntry.graphicalVoiceEntries) {
                this.setBeamNotesWantedStemDirections(voiceEntry.parentVoiceEntry);
            }
        }
        // if there are no staffEntries in this measure, create a rest for the whole measure:
        // check OSMDOptions.fillEmptyMeasuresWithWholeRest
        if (this.rules.FillEmptyMeasuresWithWholeRest >= 1) { // fill measures with no notes given with whole rests, visible (1) or invisible (2)
            if (measure.staffEntries.length === 0) {
                const sourceStaffEntry: SourceStaffEntry = new SourceStaffEntry(
                    new VerticalSourceStaffEntryContainer(measure.parentSourceMeasure,
                                                          measure.parentSourceMeasure.AbsoluteTimestamp,
                                                          measure.parentSourceMeasure.CompleteNumberOfStaves),
                    staff);
                if (staff.Voices.length === 0) {
                    const newVoice: Voice = new Voice(measure.ParentStaff.ParentInstrument, -1);
                    // this is problematic because we don't know the MusicXML voice ids and how many voices with which ids will be created after this.
                    //   but it should only happen when the first measure of the piece is empty.
                    staff.Voices.push(newVoice);
                }
                const voiceEntry: VoiceEntry = new VoiceEntry(new Fraction(0, 1), staff.Voices[0], sourceStaffEntry);
                let duration: Fraction = sourceMeasure.Duration;
                if (duration.RealValue === 0) {
                    duration = sourceMeasure.ActiveTimeSignature.clone();
                }
                const note: Note = new Note(voiceEntry, sourceStaffEntry, duration, undefined, sourceMeasure, true);
                note.IsWholeMeasureRest = true; // there may be a more elegant solution
                note.PrintObject = this.rules.FillEmptyMeasuresWithWholeRest === FillEmptyMeasuresWithWholeRests.YesVisible;
                  // don't display whole rest that wasn't given in XML, only for layout/voice completion
                voiceEntry.Notes.push(note);
                const graphicalStaffEntry: GraphicalStaffEntry = MusicSheetCalculator.symbolFactory.createStaffEntry(sourceStaffEntry, measure);
                measure.addGraphicalStaffEntry(graphicalStaffEntry);
                graphicalStaffEntry.relInMeasureTimestamp = voiceEntry.Timestamp;
                const gve: GraphicalVoiceEntry = MusicSheetCalculator.symbolFactory.createVoiceEntry(voiceEntry, graphicalStaffEntry);
                graphicalStaffEntry.graphicalVoiceEntries.push(gve);
                const graphicalNote: GraphicalNote = MusicSheetCalculator.symbolFactory.createNote(
                    note,
                    gve,
                    new ClefInstruction(),
                    OctaveEnum.NONE,
                    this.rules);
                MusicSheetCalculator.stafflineNoteCalculator.trackNote(graphicalNote);
                gve.notes.push(graphicalNote);
            }
        }

        measure.hasOnlyRests = true;
        //if staff entries empty, loop will not start. so true is valid
        for (const graphicalStaffEntry of measure.staffEntries) {
            //Loop until we get just one false
            measure.hasOnlyRests = graphicalStaffEntry.hasOnlyRests();
            if (!measure.hasOnlyRests) {
                break;
            }
        }

        return measure;
    }

    private checkNoteForAccidental(graphicalNote: GraphicalNote, accidentalCalculator: AccidentalCalculator, activeClef: ClefInstruction,
                                   octaveEnum: OctaveEnum): void {
        let pitch: Pitch = graphicalNote.sourceNote.Pitch;
        const transposeHalftones: number = graphicalNote.parentVoiceEntry.parentStaffEntry.parentMeasure.getTransposedHalftones();
        if (transposeHalftones !== 0 && graphicalNote.sourceNote.ParentStaffEntry.ParentStaff.ParentInstrument.MidiInstrumentId !== MidiInstrument.Percussion) {
            pitch = graphicalNote.Transpose(
                accidentalCalculator.ActiveKeyInstruction, activeClef, transposeHalftones, octaveEnum
            );
            graphicalNote.sourceNote.TransposedPitch = pitch;
        } else {
            // Clear any previously set TransposedPitch when not transposing,
            // to avoid stale state affecting accidental calculation (#1383)
            graphicalNote.sourceNote.TransposedPitch = undefined;
        }
        graphicalNote.sourceNote.halfTone = pitch.getHalfTone();
        accidentalCalculator.checkAccidental(graphicalNote, pitch);
    }

    // private createStaffEntryForTieNote(measure: StaffMeasure, absoluteTimestamp: Fraction, openTie: Tie): GraphicalStaffEntry {
    //     let graphicalStaffEntry: GraphicalStaffEntry;
    //     graphicalStaffEntry = MusicSheetCalculator.symbolFactory.createStaffEntry(openTie.Start.ParentStaffEntry, measure);
    //     graphicalStaffEntry.relInMeasureTimestamp = Fraction.minus(absoluteTimestamp, measure.parentSourceMeasure.AbsoluteTimestamp);
    //     this.resetYPositionForLeadSheet(graphicalStaffEntry.PositionAndShape);
    //     measure.addGraphicalStaffEntryAtTimestamp(graphicalStaffEntry);
    //     return graphicalStaffEntry;
    // }

    private handleStaffEntries(staffIsPercussionArray: Array<boolean>): void {
        for (let idx: number = 0, len: number = this.graphicalMusicSheet.MeasureList.length; idx < len; ++idx) {
            const measures: GraphicalMeasure[] = this.graphicalMusicSheet.MeasureList[idx];
            for (let idx2: number = 0, len2: number = measures.length; idx2 < len2; ++idx2) {
                const measure: GraphicalMeasure = measures[idx2];
                if (!measure) {
                    continue;
                }
                //This property is active...
                if (this.rules.PercussionOneLineCutoff > 0 && !this.rules.PercussionUseCajon2NoteSystem) {
                    //We have a percussion clef, check to see if this property applies...
                    if (staffIsPercussionArray[idx2]) {
                        //-1 means always trigger, or we are under the cutoff number specified
                        if (this.rules.PercussionOneLineCutoff === -1 ||
                            MusicSheetCalculator.stafflineNoteCalculator.getStafflineUniquePositionCount(idx2) < this.rules.PercussionOneLineCutoff) {
                            measure.ParentStaff.StafflineCount = 1;
                        }
                    }
                }
                for (const graphicalStaffEntry of measure.staffEntries) {
                    if (graphicalStaffEntry.parentMeasure !== undefined
                        && graphicalStaffEntry.graphicalVoiceEntries.length > 0
                        && graphicalStaffEntry.graphicalVoiceEntries[0].notes.length > 0) {
                        this.layoutVoiceEntries(graphicalStaffEntry, idx2);
                        this.layoutStaffEntry(graphicalStaffEntry);
                    }
                }
                this.graphicalMeasureCreatedCalculations(measure);
            }
        }
    }

    protected calculateSkyBottomLines(): void {
        // override
    }

    /**
     * Re-adjust the x positioning of expressions.
     */
    protected calculateExpressionAlignements(): void {
        // override
    }

    // does nothing for now, because layoutBeams() is an empty method
    // private calculateBeams(): void {
    //     for (let idx2: number = 0, len2: number = this.musicSystems.length; idx2 < len2; ++idx2) {
    //         const musicSystem: MusicSystem = this.musicSystems[idx2];
    //         for (let idx3: number = 0, len3: number = musicSystem.StaffLines.length; idx3 < len3; ++idx3) {
    //             const staffLine: StaffLine = musicSystem.StaffLines[idx3];
    //             for (let idx4: number = 0, len4: number = staffLine.Measures.length; idx4 < len4; ++idx4) {
    //                 const measure: GraphicalMeasure = staffLine.Measures[idx4];
    //                 for (let idx5: number = 0, len5: number = measure.staffEntries.length; idx5 < len5; ++idx5) {
    //                     const staffEntry: GraphicalStaffEntry = measure.staffEntries[idx5];
    //                     this.layoutBeams(staffEntry);
    //                 }
    //             }
    //         }
    //     }
    // }

    private calculateStaffEntryArticulationMarks(): void {
        for (let idx2: number = 0, len2: number = this.musicSystems.length; idx2 < len2; ++idx2) {
            const system: MusicSystem = this.musicSystems[idx2];
            for (let idx3: number = 0, len3: number = system.StaffLines.length; idx3 < len3; ++idx3) {
                const line: StaffLine = system.StaffLines[idx3];
                for (let idx4: number = 0, len4: number = line.Measures.length; idx4 < len4; ++idx4) {
                    const measure: GraphicalMeasure = line.Measures[idx4];
                    for (let idx5: number = 0, len5: number = measure.staffEntries.length; idx5 < len5; ++idx5) {
                        const graphicalStaffEntry: GraphicalStaffEntry = measure.staffEntries[idx5];
                        for (let idx6: number = 0, len6: number = graphicalStaffEntry.sourceStaffEntry.VoiceEntries.length; idx6 < len6; ++idx6) {
                            const voiceEntry: VoiceEntry = graphicalStaffEntry.sourceStaffEntry.VoiceEntries[idx6];
                            if (voiceEntry.Articulations.length > 0) {
                                this.layoutArticulationMarks(voiceEntry.Articulations, voiceEntry, graphicalStaffEntry);
                            }
                        }
                    }
                }
            }
        }
    }

    private calculateOrnaments(): void {
        for (let idx2: number = 0, len2: number = this.musicSystems.length; idx2 < len2; ++idx2) {
            const system: MusicSystem = this.musicSystems[idx2];
            for (let idx3: number = 0, len3: number = system.StaffLines.length; idx3 < len3; ++idx3) {
                const line: StaffLine = system.StaffLines[idx3];
                for (let idx4: number = 0, len4: number = line.Measures.length; idx4 < len4; ++idx4) {
                    const measure: GraphicalMeasure = line.Measures[idx4];
                    for (let idx5: number = 0, len5: number = measure.staffEntries.length; idx5 < len5; ++idx5) {
                        const graphicalStaffEntry: GraphicalStaffEntry = measure.staffEntries[idx5];
                        for (let idx6: number = 0, len6: number = graphicalStaffEntry.sourceStaffEntry.VoiceEntries.length; idx6 < len6; ++idx6) {
                            const voiceEntry: VoiceEntry = graphicalStaffEntry.sourceStaffEntry.VoiceEntries[idx6];
                            if (voiceEntry.OrnamentContainer) {
                                if (voiceEntry.hasTie() && !graphicalStaffEntry.relInMeasureTimestamp.Equals(voiceEntry.Timestamp)) {
                                    continue;
                                }
                                this.layoutOrnament(voiceEntry.OrnamentContainer, voiceEntry, graphicalStaffEntry);
                                if (!(this.staffEntriesWithOrnaments.indexOf(graphicalStaffEntry) !== -1)) {
                                    this.staffEntriesWithOrnaments.push(graphicalStaffEntry);
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    private getFingeringPlacement(measure: GraphicalMeasure, fingering?: TechnicalInstruction): PlacementEnum {
        let placement: PlacementEnum = this.rules.FingeringPosition;
        if (this.rules.FingeringPositionFromXML &&
            placement !== PlacementEnum.NotYetDefined &&
            placement !== PlacementEnum.AboveOrBelow &&
            (fingering?.placement === PlacementEnum.Above || fingering?.placement === PlacementEnum.Below)) {
            placement = fingering.placement;
        }
        if (placement === PlacementEnum.NotYetDefined || placement === PlacementEnum.AboveOrBelow) {
            placement = measure.isUpperStaffOfInstrument() ? PlacementEnum.Above : PlacementEnum.Below;
        }
        return placement;
    }

    private getFingeringAnchorY(placement: PlacementEnum): number {
        return this.getFingeringStaffLaneY(placement);
    }

    private getFingeringStaffLaneY(placement: PlacementEnum): number {
        if (this.rules.FingeringInsideStafflines) {
            return placement === PlacementEnum.Above ? 0 : this.rules.StaffHeight;
        }
        return placement === PlacementEnum.Above ? -0.22 : this.rules.StaffHeight + 0.22;
    }

    private getRelativePositionToAncestor(shape: BoundingBox, ancestor: BoundingBox): PointF2D {
        let x: number = shape.RelativePosition.x;
        let y: number = shape.RelativePosition.y;
        let parent: BoundingBox = shape.Parent;
        while (parent && parent !== ancestor) {
            x += parent.RelativePosition.x;
            y += parent.RelativePosition.y;
            parent = parent.Parent;
        }
        return parent === ancestor ? new PointF2D(x, y) : undefined;
    }

    private getFingeringOwnerNote(gse: GraphicalStaffEntry, fingering: TechnicalInstruction): GraphicalNote {
        let fallback: GraphicalNote = undefined;
        for (const voiceEntry of gse.graphicalVoiceEntries) {
            const ownsInstruction: boolean = voiceEntry.parentVoiceEntry?.TechnicalInstructions?.indexOf(fingering) >= 0;
            for (const graphicalNote of voiceEntry.notes) {
                const sourceNote: Note = graphicalNote.sourceNote;
                if (sourceNote === fingering.sourceNote || sourceNote.Fingering === fingering) {
                    return graphicalNote;
                }
                if (!fallback && ownsInstruction) {
                    fallback = graphicalNote;
                }
            }
        }
        return fallback;
    }

    private getFingeringLabelHeight(label: GraphicalLabel): number {
        label.setLabelPositionAndShapeBorders();
        return Math.max(0.1, label.PositionAndShape.BorderMarginBottom - label.PositionAndShape.BorderMarginTop);
    }

    private setFingeringLabelPosition(label: GraphicalLabel, x: number, y: number, placement: PlacementEnum): void {
        label.Label.textAlignment = placement === PlacementEnum.Above ?
            TextAlignmentEnum.CenterBottom : TextAlignmentEnum.CenterTop;
        label.setLabelPositionAndShapeBorders();
        label.PositionAndShape.RelativePosition = new PointF2D(x, y);
        label.PositionAndShape.calculateBoundingBox();
    }

    private getFingeringStackSpacing(labelHeight: number): number {
        return Math.max(this.rules.FingeringPaddingY + labelHeight * 0.58, 0.42);
    }

    private getFingeringLaneOrder(gse: GraphicalStaffEntry, fingeringOwners: GraphicalNote[]): GraphicalNote[] {
        const laneNotes: GraphicalNote[] = [];
        for (const owner of fingeringOwners) {
            if (owner && laneNotes.indexOf(owner) < 0) {
                laneNotes.push(owner);
            }
        }

        const voicedEntries: GraphicalVoiceEntry[] = gse.graphicalVoiceEntries.filter(
            (voiceEntry: GraphicalVoiceEntry) => !voiceEntry.parentVoiceEntry?.IsGrace && voiceEntry.notes.length > 0);
        if (voicedEntries.length > 1) {
            for (const voiceEntry of voicedEntries) {
                const sortedNotes: GraphicalNote[] = voiceEntry.notes.slice().sort(
                    (a: GraphicalNote, b: GraphicalNote) => this.getGraphicalNoteCenterY(a) - this.getGraphicalNoteCenterY(b));
                const topNote: GraphicalNote = sortedNotes[0];
                const bottomNote: GraphicalNote = sortedNotes[sortedNotes.length - 1];
                if (topNote && laneNotes.indexOf(topNote) < 0) {
                    laneNotes.push(topNote);
                }
                if (bottomNote && laneNotes.indexOf(bottomNote) < 0) {
                    laneNotes.push(bottomNote);
                }
            }
        }

        return laneNotes.sort((a: GraphicalNote, b: GraphicalNote) => this.getGraphicalNoteCenterY(a) - this.getGraphicalNoteCenterY(b));
    }

    private getGraphicalNoteCenterY(note: GraphicalNote): number {
        if (Number.isFinite(note?.staffLine)) {
            return note.staffLine;
        }
        const noteY: number = note?.PositionAndShape?.RelativePosition?.y;
        if (Number.isFinite(noteY)) {
            return noteY;
        }
        return 0;
    }

    private getFingeringStackSlots(fingerings: TechnicalInstruction[], owners: GraphicalNote[],
                                   gse: GraphicalStaffEntry, placement: PlacementEnum): number[] {
        const laneOrder: GraphicalNote[] = this.getFingeringLaneOrder(gse, owners);
        const baseSlotOccurrences: Map<number, number> = new Map<number, number>();
        let nextFallbackSlot: number = Math.min(laneOrder.length, 1);

        return fingerings.map((_fingering: TechnicalInstruction, index: number) => {
            const owner: GraphicalNote = owners[index];
            if (!owner || laneOrder.length === 0) {
                return nextFallbackSlot++;
            }

            const laneIndex: number = laneOrder.indexOf(owner);
            const distanceFromPlacementSide: number = placement === PlacementEnum.Above ?
                laneIndex : laneOrder.length - 1 - laneIndex;
            const baseSlot: number = Math.min(distanceFromPlacementSide, 1);
            const occurrence: number = baseSlotOccurrences.get(baseSlot) ?? 0;
            baseSlotOccurrences.set(baseSlot, occurrence + 1);
            return baseSlot + occurrence;
        });
    }

    private getBoundingBoxRectInStaffLine(shape: BoundingBox, line: StaffLine, margin: boolean = false): FingeringCollisionRect {
        const relativePosition: PointF2D = this.getRelativePositionToAncestor(shape, line.PositionAndShape);
        if (!relativePosition) {
            return undefined;
        }
        return {
            left: relativePosition.x + (margin ? shape.BorderMarginLeft : shape.BorderLeft),
            right: relativePosition.x + (margin ? shape.BorderMarginRight : shape.BorderRight),
            top: relativePosition.y + (margin ? shape.BorderMarginTop : shape.BorderTop),
            bottom: relativePosition.y + (margin ? shape.BorderMarginBottom : shape.BorderBottom),
            penalty: 1
        };
    }

    private getMeasureNumberRectInStaffLine(label: GraphicalLabel, line: StaffLine): FingeringCollisionRect {
        const labelPosition: PointF2D = label.PositionAndShape.RelativePosition;
        const staffLinePosition: PointF2D = line.PositionAndShape.RelativePosition;
        return {
            left: labelPosition.x - staffLinePosition.x + label.PositionAndShape.BorderMarginLeft,
            right: labelPosition.x - staffLinePosition.x + label.PositionAndShape.BorderMarginRight,
            top: labelPosition.y - staffLinePosition.y + label.PositionAndShape.BorderMarginTop,
            bottom: labelPosition.y - staffLinePosition.y + label.PositionAndShape.BorderMarginBottom,
            penalty: 1.5,
            kind: "measure-number"
        };
    }

    private addRectIfValid(rects: FingeringCollisionRect[], rect: FingeringCollisionRect): void {
        if (!rect ||
            !Number.isFinite(rect.left) || !Number.isFinite(rect.right) ||
            !Number.isFinite(rect.top) || !Number.isFinite(rect.bottom) ||
            rect.right <= rect.left || rect.bottom <= rect.top) {
            return;
        }
        rects.push(rect);
    }

    private withFingeringDebugSource(rect: FingeringCollisionRect, measure: GraphicalMeasure,
                                     line: StaffLine, debugKey?: string): FingeringCollisionRect {
        if (!rect) {
            return rect;
        }
        rect.debugMeasure = measure?.MeasureNumber;
        rect.debugMeasureX = measure?.PositionAndShape?.RelativePosition?.x;
        rect.debugStaffId = line?.ParentStaff?.Id;
        rect.debugKey = debugKey ?? rect.debugKey;
        return rect;
    }

    private expandFingeringRect(rect: FingeringCollisionRect, padding: number): FingeringCollisionRect {
        return {
            left: rect.left - padding,
            right: rect.right + padding,
            top: rect.top - padding,
            bottom: rect.bottom + padding,
            penalty: rect.penalty,
            maxShift: rect.maxShift,
            soft: rect.soft,
            isFingering: rect.isFingering,
            kind: rect.kind,
            debugMeasure: rect.debugMeasure,
            debugMeasureX: rect.debugMeasureX,
            debugStaffId: rect.debugStaffId,
            debugKey: rect.debugKey
        };
    }

    private shiftFingeringRect(rect: FingeringCollisionRect, dx: number, dy: number): FingeringCollisionRect {
        return {
            left: rect.left + dx,
            right: rect.right + dx,
            top: rect.top + dy,
            bottom: rect.bottom + dy,
            penalty: rect.penalty,
            maxShift: rect.maxShift,
            soft: rect.soft,
            isFingering: rect.isFingering,
            kind: rect.kind,
            debugMeasure: rect.debugMeasure,
            debugMeasureX: rect.debugMeasureX,
            debugStaffId: rect.debugStaffId,
            debugKey: rect.debugKey
        };
    }

    private getRectOverlapArea(a: FingeringCollisionRect, b: FingeringCollisionRect): number {
        const overlapWidth: number = Math.min(a.right, b.right) - Math.max(a.left, b.left);
        const overlapHeight: number = Math.min(a.bottom, b.bottom) - Math.max(a.top, b.top);
        if (overlapWidth <= 0 || overlapHeight <= 0) {
            return 0;
        }
        return overlapWidth * overlapHeight;
    }

    private getLabelRect(label: GraphicalLabel): FingeringCollisionRect {
        return {
            left: label.PositionAndShape.RelativePosition.x + label.PositionAndShape.BorderMarginLeft,
            right: label.PositionAndShape.RelativePosition.x + label.PositionAndShape.BorderMarginRight,
            top: label.PositionAndShape.RelativePosition.y + label.PositionAndShape.BorderMarginTop,
            bottom: label.PositionAndShape.RelativePosition.y + label.PositionAndShape.BorderMarginBottom,
            penalty: 1
        };
    }

    private getMeasureRectInStaffLine(measure: GraphicalMeasure, line: StaffLine): FingeringCollisionRect {
        const relativePosition: PointF2D = this.getRelativePositionToAncestor(measure.PositionAndShape, line.PositionAndShape);
        if (!relativePosition) {
            return undefined;
        }
        const left: number = relativePosition.x;
        return {
            left,
            right: left + measure.PositionAndShape.Size.width,
            top: relativePosition.y + measure.PositionAndShape.BorderTop,
            bottom: relativePosition.y + measure.PositionAndShape.BorderBottom,
            penalty: 1
        };
    }

    private snapFingeringXToMeasure(label: GraphicalLabel, desiredX: number, measure: GraphicalMeasure,
                                    line: StaffLine): number {
        const measureRect: FingeringCollisionRect = this.getMeasureRectInStaffLine(measure, line);
        if (!measureRect) {
            return desiredX;
        }

        const inset: number = 0.06;
        const leftLimit: number = measureRect.left + inset;
        const rightLimit: number = measureRect.right - inset;
        const labelLeftOffset: number = label.PositionAndShape.BorderMarginLeft;
        const labelRightOffset: number = label.PositionAndShape.BorderMarginRight;
        const labelWidth: number = labelRightOffset - labelLeftOffset;
        const measureWidth: number = rightLimit - leftLimit;
        if (measureWidth <= 0 || labelWidth <= 0) {
            return desiredX;
        }
        if (labelWidth >= measureWidth) {
            return (leftLimit + rightLimit - labelLeftOffset - labelRightOffset) / 2;
        }
        return Math.min(
            Math.max(desiredX, leftLimit - labelLeftOffset),
            rightLimit - labelRightOffset
        );
    }

    private getCompactFingeringStackRect(rect: FingeringCollisionRect): FingeringCollisionRect {
        const height: number = rect.bottom - rect.top;
        const reservedHeight: number = Math.min(height, 0.62);
        const verticalInset: number = Math.max(0, (height - reservedHeight) / 2);
        return {
            left: rect.left,
            right: rect.right,
            top: rect.top + verticalInset,
            bottom: rect.bottom - verticalInset,
            penalty: rect.penalty,
            maxShift: rect.maxShift,
            soft: rect.soft,
            isFingering: rect.isFingering,
            kind: rect.kind,
            debugMeasure: rect.debugMeasure,
            debugMeasureX: rect.debugMeasureX,
            debugStaffId: rect.debugStaffId,
            debugKey: rect.debugKey
        };
    }

    private getExistingFingeringCollisionRect(label: GraphicalLabel, placement: PlacementEnum): FingeringCollisionRect {
        const rect: FingeringCollisionRect = this.getLabelRect(label);
        rect.isFingering = true;
        rect.kind = "existing-fingering";
        return placement === PlacementEnum.Below ? this.getCompactFingeringStackRect(rect) : rect;
    }

    private addFingeringNoteCollisionRects(rects: FingeringCollisionRect[], graphicalNote: GraphicalNote,
                                           line: StaffLine, measure: GraphicalMeasure): void {
        const noteRect: FingeringCollisionRect = this.getBoundingBoxRectInStaffLine(graphicalNote.PositionAndShape, line, true);
        if (noteRect) {
            noteRect.kind = "note";
        }
        const noteId: string = (graphicalNote as any)?.getSVGId?.();
        this.addRectIfValid(rects, this.withFingeringDebugSource(noteRect, measure, line, noteId ? `note:${noteId}` : undefined));
        this.addFingeringNoteheadCollisionRect(rects, graphicalNote, measure, line);
        this.addFingeringStemCollisionRect(rects, graphicalNote, measure, line);
    }

    private addFingeringRestCollisionRects(rects: FingeringCollisionRect[], restNote: GraphicalNote,
                                           line: StaffLine, measure: GraphicalMeasure): void {
        const restRects: FingeringCollisionRect[] = this.getRestCollisionRects(restNote, measure, line);
        const restId: string = (restNote as any)?.getSVGId?.();
        for (let restRectIndex: number = 0; restRectIndex < restRects.length; restRectIndex++) {
            this.addRectIfValid(rects, this.withFingeringDebugSource(
                restRects[restRectIndex],
                measure,
                line,
                restId ? `rest:${restId}:${restRectIndex}` : `rest:${restRectIndex}`
            ));
        }
    }

    private addFingeringNoteheadCollisionRect(rects: FingeringCollisionRect[], graphicalNote: GraphicalNote,
                                              measure: GraphicalMeasure, line: StaffLine): void {
        this.addRectIfValid(rects, this.getFingeringNoteheadCollisionRect(graphicalNote, measure, line));
    }

    private getFingeringNoteheadCollisionRect(graphicalNote: GraphicalNote, measure: GraphicalMeasure,
                                              line: StaffLine): FingeringCollisionRect {
        if (!graphicalNote || graphicalNote.sourceNote?.isRest() ||
            graphicalNote.sourceNote?.Notehead?.Shape === NoteHeadShape.NONE) {
                return undefined;
        }
        const vfNote: any = (graphicalNote as any).vfnote?.[0];
        const noteheadIndex: number = Number((graphicalNote as any).vfnote?.[1] ?? (graphicalNote as any).vfnoteIndex ?? 0);
        if (!vfNote || !Number.isFinite(noteheadIndex) || typeof vfNote.getYs !== "function") {
            return undefined;
        }
        const ys: number[] = vfNote.getYs();
        if (!Array.isArray(ys) || !Number.isFinite(ys[noteheadIndex])) {
            return undefined;
        }
        const xs: number[] = typeof vfNote.getXs === "function" ? vfNote.getXs() : [];
        const fallbackX: number = typeof vfNote.getStemX === "function" ? vfNote.getStemX() : undefined;
        const x: number = Number.isFinite(xs[noteheadIndex]) ? xs[noteheadIndex] : xs.length > 0 ? xs[xs.length - 1] : fallbackX;
        if (!Number.isFinite(x)) {
            return undefined;
        }
        const radius: number = this.estimateFingeringNoteheadRadiusPx(vfNote);
        const staveOrigin: { x: number, y: number } = this.getVexFlowStaveOriginPx(vfNote);
        const measureX: number = measure.PositionAndShape.RelativePosition.x;
        const rect: FingeringCollisionRect = this.fingeringPxRectToCollisionRect({
            x: x - radius,
            y: ys[noteheadIndex] - radius,
            width: radius * 2,
            height: radius * 2
        }, measureX, staveOrigin, "notehead", 1);
        if (!rect || rect.bottom < -4 || rect.top > this.rules.StaffHeight + 4) {
            return undefined;
        }
        const noteId: string = (graphicalNote as any)?.getSVGId?.();
        return this.withFingeringDebugSource(
            rect,
            measure,
            line,
            noteId ? `notehead:${noteId}:${noteheadIndex}` : `notehead:${noteheadIndex}`
        );
    }

    private estimateFingeringNoteheadRadiusPx(vfNote: any): number {
        const glyphWidth: number = typeof vfNote?.getGlyphWidth === "function" ? vfNote.getGlyphWidth() : 10;
        if (Number.isFinite(glyphWidth) && glyphWidth > 0) {
            return Math.max(3, Math.min(8, glyphWidth * 0.42));
        }
        return 5;
    }

    private addFingeringStemCollisionRect(rects: FingeringCollisionRect[], graphicalNote: GraphicalNote,
                                          measure: GraphicalMeasure, line: StaffLine): void {
        if (!graphicalNote || graphicalNote.sourceNote?.isRest()) {
            return;
        }
        const vfNote: any = (graphicalNote as any).vfnote?.[0];
        if (!vfNote || typeof vfNote.getStemX !== "function" || typeof vfNote.getStemExtents !== "function") {
            return;
        }
        let stemX: number;
        let topY: number;
        let bottomY: number;
        try {
            stemX = vfNote.getStemX();
            const extents: any = vfNote.getStemExtents();
            topY = extents?.topY;
            bottomY = extents?.baseY;
        } catch (e) {
            return;
        }
        if (!Number.isFinite(stemX) || !Number.isFinite(topY) || !Number.isFinite(bottomY)) {
            return;
        }
        const staveOrigin: { x: number, y: number } = this.getVexFlowStaveOriginPx(vfNote);
        const measureX: number = graphicalNote.parentVoiceEntry?.parentStaffEntry?.parentMeasure?.PositionAndShape?.RelativePosition?.x ?? 0;
        const stemWidth: number = Math.max(this.rules.StemWidth, 0.08);
        const stemStaffX: number = measureX + (stemX - staveOrigin.x) / 10;
        const top: number = (Math.min(topY, bottomY) - staveOrigin.y) / 10;
        const bottom: number = (Math.max(topY, bottomY) - staveOrigin.y) / 10;
        if (bottom < -4 || top > this.rules.StaffHeight + 4) {
            return;
        }
        const noteId: string = (graphicalNote as any)?.getSVGId?.();
        this.addRectIfValid(rects, this.withFingeringDebugSource({
            left: stemStaffX - stemWidth / 2,
            right: stemStaffX + stemWidth / 2,
            top,
            bottom,
            penalty: 1,
            kind: "stem"
        }, measure, line, noteId ? `stem:${noteId}` : undefined));
    }

    private getVexFlowStaveOriginPx(vfElement: any): { x: number, y: number } {
        const stave: any = vfElement?.getStave?.() ?? vfElement?.stave;
        const rawX: number = typeof stave?.getX === "function" ? stave.getX() : stave?.x;
        const rawY: number = typeof stave?.getY === "function" ? stave.getY() : stave?.y;
        return {
            x: Number.isFinite(rawX) ? rawX : 0,
            y: Number.isFinite(rawY) ? rawY : 0
        };
    }

    private normalizeFingeringPxRect(rect: FingeringPxRect, padding: number = 0): FingeringPxRect {
        const x1: number = Math.min(rect.x, rect.x + rect.width) - padding;
        const x2: number = Math.max(rect.x, rect.x + rect.width) + padding;
        const y1: number = Math.min(rect.y, rect.y + rect.height) - padding;
        const y2: number = Math.max(rect.y, rect.y + rect.height) + padding;
        if (!Number.isFinite(x1) || !Number.isFinite(x2) || !Number.isFinite(y1) || !Number.isFinite(y2)) {
            return undefined;
        }
        return {
            x: x1,
            y: y1,
            width: x2 - x1,
            height: y2 - y1
        };
    }

    private fingeringPxRectToCollisionRect(rect: FingeringPxRect, measureX: number,
                                           staveOrigin: { x: number, y: number },
                                           kind: string, penalty: number): FingeringCollisionRect {
        if (!rect) {
            return undefined;
        }
        return {
            left: measureX + (rect.x - staveOrigin.x) / 10,
            right: measureX + (rect.x + rect.width - staveOrigin.x) / 10,
            top: (rect.y - staveOrigin.y) / 10,
            bottom: (rect.y + rect.height - staveOrigin.y) / 10,
            penalty,
            kind
        };
    }

    private getVexFlowElementCategory(element: any): string {
        const category: any = typeof element?.getCategory === "function" ? element.getCategory() : element?.category;
        return `${category ?? ""}`.toLowerCase();
    }

    private getVexFlowNoteModifiers(vfNote: any): any[] {
        const modifiers: any[] = typeof vfNote?.getModifiers === "function" ? vfNote.getModifiers() : vfNote?.modifiers;
        return Array.isArray(modifiers) ? modifiers : [];
    }

    private addFingeringAttachedModifierCollisionRects(rects: FingeringCollisionRect[], gse: GraphicalStaffEntry,
                                                       line: StaffLine, measure: GraphicalMeasure): Set<string> {
        const addedKinds: Set<string> = new Set<string>();
        const seenModifiers: Set<any> = new Set<any>();
        const measureX: number = measure.PositionAndShape.RelativePosition.x;
        for (let voiceEntryIndex: number = 0; voiceEntryIndex < gse.graphicalVoiceEntries.length; voiceEntryIndex++) {
            const voiceEntry: GraphicalVoiceEntry = gse.graphicalVoiceEntries[voiceEntryIndex];
            for (let noteIndex: number = 0; noteIndex < voiceEntry.notes.length; noteIndex++) {
                const graphicalNote: GraphicalNote = voiceEntry.notes[noteIndex];
                const vfNote: any = (graphicalNote as any).vfnote?.[0] ?? (voiceEntry as any).vfStaveNote;
                const modifiers: any[] = this.getVexFlowNoteModifiers(vfNote);
                if (!vfNote || modifiers.length === 0) {
                    continue;
                }
                const staveOrigin: { x: number, y: number } = this.getVexFlowStaveOriginPx(vfNote);
                for (let modifierIndex: number = 0; modifierIndex < modifiers.length; modifierIndex++) {
                    const modifier: any = modifiers[modifierIndex];
                    if (seenModifiers.has(modifier)) {
                        continue;
                    }
                    seenModifiers.add(modifier);
                    const category: string = this.getVexFlowElementCategory(modifier);
                    let kind: string;
                    let pxRect: FingeringPxRect;
                    if (category.indexOf("ornament") >= 0) {
                        kind = "ornament";
                        pxRect = this.getFingeringOrnamentCollisionRectPx(modifier);
                    } else if (category.indexOf("articulation") >= 0) {
                        kind = "articulation";
                        pxRect = this.getFingeringArticulationCollisionRectPx(modifier);
                    } else {
                        continue;
                    }
                    const rect: FingeringCollisionRect =
                        this.fingeringPxRectToCollisionRect(pxRect, measureX, staveOrigin, kind, 1);
                    if (!rect || rect.bottom < -4 || rect.top > this.rules.StaffHeight + 4) {
                        continue;
                    }
                    this.addRectIfValid(rects, this.withFingeringDebugSource(
                        rect,
                        measure,
                        line,
                        `${kind}:${voiceEntryIndex}:${noteIndex}:${modifierIndex}`
                    ));
                    addedKinds.add(kind);
                }
            }
        }
        return addedKinds;
    }

    private getFingeringModifierPositions(): any {
        return (VF.Modifier as any)?.Position ?? { LEFT: 1, RIGHT: 2, ABOVE: 3, BELOW: 4 };
    }

    private getFingeringModifierPosition(modifier: any): number {
        return Number(typeof modifier?.getPosition === "function" ? modifier.getPosition() : modifier?.position);
    }

    private getFingeringVexFlowStemUp(): number {
        return Number((VF.Stem as any)?.UP ?? 1);
    }

    private getFingeringVexFlowStemDown(): number {
        return Number((VF.Stem as any)?.DOWN ?? -1);
    }

    private getFingeringNoteStemDirection(note: any): number {
        const direction: number = Number(typeof note?.getStemDirection === "function" ? note.getStemDirection() : note?.stem_direction);
        return Number.isFinite(direction) ? direction : this.getFingeringVexFlowStemUp();
    }

    private getFingeringNoteStemExtents(note: any): { topY?: number, baseY?: number } {
        try {
            if (typeof note?.getStemExtents === "function") {
                return note.getStemExtents();
            }
            const stem: any = typeof note?.getStem === "function" ? note.getStem() : note?.stem;
            if (typeof stem?.getExtents === "function") {
                return stem.getExtents();
            }
        } catch (e) {
            return {};
        }
        return {};
    }

    private getFingeringAttachedModifierTopYPx(note: any, textLine: number): number {
        const stave: any = typeof note?.getStave === "function" ? note.getStave() : note?.stave;
        const stemDirection: number = this.getFingeringNoteStemDirection(note);
        const stemExtents: any = this.getFingeringNoteStemExtents(note);
        if (note.getCategory?.() === "tabnotes") {
            if (typeof note.hasStem === "function" && note.hasStem() &&
                stemDirection === this.getFingeringVexFlowStemUp()) {
                    return Number(stemExtents?.topY);
            }
            return Number(stave?.getYForTopText?.(textLine));
        }
        if (typeof note.hasStem === "function" && note.hasStem()) {
            return stemDirection === this.getFingeringVexFlowStemUp()
                ? Number(stemExtents?.topY)
                : Number(stemExtents?.baseY);
        }
        const ys: number[] = typeof note.getYs === "function" ? note.getYs() : [];
        const finiteYs: number[] = Array.isArray(ys) ? ys.filter((yValue: number) => Number.isFinite(yValue)) : [];
        return finiteYs.length > 0 ? Math.min(...finiteYs) : undefined;
    }

    private getFingeringAttachedModifierBottomYPx(note: any, textLine: number): number {
        const stave: any = typeof note?.getStave === "function" ? note.getStave() : note?.stave;
        const stemDirection: number = this.getFingeringNoteStemDirection(note);
        const stemExtents: any = this.getFingeringNoteStemExtents(note);
        if (note.getCategory?.() === "tabnotes") {
            if (typeof note.hasStem === "function" && note.hasStem() &&
                stemDirection === this.getFingeringVexFlowStemDown()) {
                    return Number(stemExtents?.topY);
            }
            return Number(stave?.getYForBottomText?.(textLine));
        }
        if (typeof note.hasStem === "function" && note.hasStem()) {
            return stemDirection === this.getFingeringVexFlowStemUp()
                ? Number(stemExtents?.baseY)
                : Number(stemExtents?.topY);
        }
        const ys: number[] = typeof note.getYs === "function" ? note.getYs() : [];
        const finiteYs: number[] = Array.isArray(ys) ? ys.filter((yValue: number) => Number.isFinite(yValue)) : [];
        return finiteYs.length > 0 ? Math.max(...finiteYs) : undefined;
    }

    private getFingeringArticulationInitialOffset(note: any, position: number): number {
        const positions: any = this.getFingeringModifierPositions();
        const isOnStemTip: boolean =
            (position === positions.ABOVE && this.getFingeringNoteStemDirection(note) === this.getFingeringVexFlowStemUp()) ||
            (position === positions.BELOW && this.getFingeringNoteStemDirection(note) === this.getFingeringVexFlowStemDown());
        if (note.getCategory?.() === "stavenotes" || note.getCategory?.() === "gracenotes") {
            return typeof note.hasStem === "function" && note.hasStem() && isOnStemTip ? 0.5 : 1;
        }
        return typeof note.hasStem === "function" && note.hasStem() && isOnStemTip ? 1 : 0;
    }

    private fingeringArticulationLineIsWithinStaff(line: number, position: number): boolean {
        const positions: any = this.getFingeringModifierPositions();
        return position === positions.ABOVE ? line <= 5 : line >= 1;
    }

    private snapFingeringArticulationLineToStaff(canSitBetweenLines: boolean, line: number,
                                                 position: number, offsetDirection: number): number {
        const roundToNearestHalf: (mathFn: (value: number) => number, value: number) => number =
            (mathFn: (value: number) => number, value: number): number => mathFn(value / 0.5) * 0.5;
        const positions: any = this.getFingeringModifierPositions();
        const roundingFunction: (value: number) => number = this.fingeringArticulationLineIsWithinStaff(line, position)
            ? position === positions.ABOVE ? Math.ceil : Math.floor
            : Math.round;
        const snappedLine: number = roundToNearestHalf(roundingFunction, line);
        const canSnapToStaffSpace: boolean =
            canSitBetweenLines && this.fingeringArticulationLineIsWithinStaff(snappedLine, position);
        return canSnapToStaffSpace && snappedLine % 1 === 0 ? snappedLine + (0.5 * -offsetDirection) : snappedLine;
    }

    private getFingeringBreathMarkXShiftPx(articulation: any, note: any, stave: any): number {
        const breathMarkDistance: number = Number(articulation.breathMarkDistance ?? 0.8);
        const noteTickContext: any = typeof note.getTickContext === "function" ? note.getTickContext() : note.tickContext;
        const nextContext: any = (VF.TickContext as any)?.getNextContext?.(noteTickContext);
        const noteX: number = Number(noteTickContext?.getX?.() ?? noteTickContext?.x ?? note?.getX?.());
        if (!Number.isFinite(noteX)) {
            return 0;
        }
        if (nextContext && Number(nextContext.x) > Number(noteTickContext?.x)) {
            return (Number(nextContext.getX?.() ?? nextContext.x) - noteX) * breathMarkDistance;
        }
        return (Number(stave.getX?.() ?? stave.x ?? 0) + Number(stave.getWidth?.() ?? stave.width ?? 0) -
            noteX + Number(stave.start_x ?? 0)) * breathMarkDistance;
    }

    private getFingeringDelayedOrnamentXShiftPx(ornament: any, note: any, stave: any): number {
        if (Number.isFinite(ornament.delayXShift)) {
            return ornament.delayXShift;
        }
        const width: number = Math.max(this.getFingeringOrnamentGlyphMetrics(ornament).width, 0);
        let delayXShift: number = width / 2;
        const noteTickContext: any = typeof note.getTickContext === "function" ? note.getTickContext() : note.tickContext;
        const nextContext: any = (VF.TickContext as any)?.getNextContext?.(noteTickContext);
        const noteX: number = Number(noteTickContext?.getX?.() ?? noteTickContext?.x ?? note?.getX?.());
        if (nextContext && Number.isFinite(noteX)) {
            delayXShift += (Number(nextContext.getX?.() ?? nextContext.x) - noteX) * 0.5;
        } else {
            const staveEnd: number = Number(stave.x ?? stave.getX?.() ?? 0) + Number(stave.width ?? stave.getWidth?.() ?? 0);
            const startX: number = Number.isFinite(noteX) ? noteX : Number(note?.getAbsoluteX?.() ?? 0);
            delayXShift += (staveEnd - startX) * 0.5;
        }
        ornament.delayXShift = delayXShift;
        return delayXShift;
    }

    private getFingeringVexflowOrnamentName(ornamentContainer: OrnamentContainer): string {
        if (ornamentContainer.VexflowOrnament) {
            return ornamentContainer.VexflowOrnament;
        }
        switch (ornamentContainer.GetOrnament) {
            case OrnamentEnum.LongTrill:
            case OrnamentEnum.Trill:
                return "tr";
            case OrnamentEnum.DelayedInvertedTurn:
            case OrnamentEnum.InvertedTurn:
                return "turn_inverted";
            case OrnamentEnum.DelayedTurn:
            case OrnamentEnum.Turn:
                return "turn";
            case OrnamentEnum.InvertedMordent:
                return "mordent";
            case OrnamentEnum.Mordent:
                return "mordent_inverted";
            case OrnamentEnum.LongMordent:
                return "prallmordent";
            case OrnamentEnum.LongInvertedMordent:
                return "tremblement";
            case OrnamentEnum.UpPrall:
                return "upprall";
            case OrnamentEnum.DownPrall:
                return "downprall";
            case OrnamentEnum.PrallUp:
                return "prallup";
            case OrnamentEnum.PrallDown:
                return "pralldown";
            case OrnamentEnum.UpMordent:
                return "upmordent";
            case OrnamentEnum.DownMordent:
                return "downmordent";
            case OrnamentEnum.LinePrall:
                return "lineprall";
            case OrnamentEnum.PrallPrall:
                return "prallprall";
            default:
                return undefined;
        }
    }

    private isFingeringDelayedOrnament(ornamentContainer: OrnamentContainer): boolean {
        return ornamentContainer.GetOrnament === OrnamentEnum.DelayedTurn ||
            ornamentContainer.GetOrnament === OrnamentEnum.DelayedInvertedTurn;
    }

    private createFingeringOrnamentEstimateModifier(ornamentContainer: OrnamentContainer, vfNote: any): any {
        const ornamentName: string = this.getFingeringVexflowOrnamentName(ornamentContainer);
        if (!ornamentName || !vfNote) {
            return undefined;
        }
        const ornament: any = new VF.Ornament(ornamentName);
        if (typeof ornament.setDelayed === "function") {
            ornament.setDelayed(this.isFingeringDelayedOrnament(ornamentContainer));
        }
        if (ornamentContainer.AccidentalBelow !== AccidentalEnum.NONE && typeof ornament.setLowerAccidental === "function") {
            ornament.setLowerAccidental(Pitch.accidentalVexflow(ornamentContainer.AccidentalBelow));
        }
        if (ornamentContainer.AccidentalAbove !== AccidentalEnum.NONE && typeof ornament.setUpperAccidental === "function") {
            ornament.setUpperAccidental(Pitch.accidentalVexflow(ornamentContainer.AccidentalAbove));
        }
        const positions: any = this.getFingeringModifierPositions();
        const position: number = ornamentContainer.placement === PlacementEnum.Below ? positions.BELOW : positions.ABOVE;
        if (typeof ornament.setPosition === "function") {
            ornament.setPosition(position);
        } else {
            ornament.position = position;
        }
        ornament.note = vfNote;
        ornament.index = 0;
        return ornament;
    }

    private getFingeringVexFlowGlyphMetrics(glyph: any, owner: any): { width: number, height: number } {
        if (!glyph) {
            return { width: 0, height: 0 };
        }
        const metrics: any = typeof glyph.getMetrics === "function" ? glyph.getMetrics() : glyph.metrics;
        const width: number = Number(metrics?.width ?? glyph.width ?? owner?.getWidth?.() ?? owner?.width ?? 0);
        const height: number = Number(metrics?.height ?? glyph.height ?? owner?.render_options?.font_scale ?? 0);
        return {
            width: Number.isFinite(width) ? width : 0,
            height: Number.isFinite(height) ? height : 0,
        };
    }

    private getFingeringOrnamentGlyphMetrics(ornament: any): FingeringOrnamentVisualMetrics {
        const glyphMetrics: { width: number, height: number } =
            this.getFingeringVexFlowGlyphMetrics(ornament?.glyph, ornament);
        if (ornament?.ornament?.smuflGlyph) {
            const fontSize: number = Number(ornament.render_options?.font_scale ?? 38) * 0.92 *
                Number(ornament.ornament?.smuflScale ?? 1);
            const safeFontSize: number = Number.isFinite(fontSize) ? fontSize : Math.max(glyphMetrics.height, 10);
            const visualHeight: number = Math.max(
                safeFontSize * 0.62,
                Math.min(glyphMetrics.height, safeFontSize) * 0.7,
                6
            );
            return {
                width: Math.max(Number(ornament?.getWidth?.() ?? ornament?.width ?? 0), glyphMetrics.width, safeFontSize * 0.7),
                height: visualHeight,
                bottomOffset: safeFontSize * 0.14,
                advanceHeight: safeFontSize,
            };
        }
        return {
            width: glyphMetrics.width,
            height: glyphMetrics.height,
            bottomOffset: 0,
            advanceHeight: glyphMetrics.height,
        };
    }

    private getFingeringArticulationCollisionRectPx(articulation: any): FingeringPxRect {
        const note: any = typeof articulation?.getNote === "function" ? articulation.getNote() : articulation?.note;
        const index: number = Number(typeof articulation?.getIndex === "function" ? articulation.getIndex() : articulation?.index);
        const position: number = this.getFingeringModifierPosition(articulation);
        const positions: any = this.getFingeringModifierPositions();
        const stave: any = typeof note?.getStave === "function" ? note.getStave() : note?.stave;
        if (!note || !stave || !Number.isFinite(index) || !Number.isFinite(position)) {
            return undefined;
        }

        const metrics: { width: number, height: number } =
            this.getFingeringVexFlowGlyphMetrics(articulation?.glyph, articulation);
        const width: number = Math.max(metrics.width, Number(articulation?.getWidth?.() ?? articulation?.width ?? 0), 4);
        const height: number = Math.max(metrics.height, 4);
        const start: any = typeof note.getModifierStartXY === "function"
            ? note.getModifierStartXY(position, index)
            : undefined;
        if (!start || !Number.isFinite(start.x)) {
            return undefined;
        }

        let x: number = start.x;
        if (articulation?.type === "abr") {
            x += this.getFingeringBreathMarkXShiftPx(articulation, note, stave);
        }
        const xShift: number = Number(typeof articulation.getXShift === "function" ?
            articulation.getXShift() : articulation.x_shift);
        if (Number.isFinite(xShift)) {
            x += xShift;
        }

        const staffSpace: number = Number(stave.getSpacingBetweenLines?.() ?? 10);
        const textLine: number = Number(articulation.text_line ?? 0);
        const canSitBetweenLines: boolean = articulation?.articulation?.between_lines === true;
        const isTab: boolean = note.getCategory?.() === "tabnotes";
        const shouldSitOutsideStaff: boolean = !canSitBetweenLines || isTab;
        const initialOffset: number = this.getFingeringArticulationInitialOffset(note, position);
        let y: number;
        if (position === positions.ABOVE) {
            const topY: number = this.getFingeringAttachedModifierTopYPx(note, textLine);
            y = topY - ((textLine + initialOffset) * staffSpace);
            if (shouldSitOutsideStaff) {
                y = Math.min(Number(stave.getYForTopText?.(-0.5) ?? y), y);
            }
        } else if (position === positions.BELOW) {
            const bottomY: number = this.getFingeringAttachedModifierBottomYPx(note, textLine);
            y = bottomY + ((textLine + initialOffset) * staffSpace);
            if (shouldSitOutsideStaff) {
                y = Math.max(Number(stave.getYForBottomText?.(-0.5) ?? y), y);
            }
        } else {
            const ys: number[] = typeof note.getYs === "function" ? note.getYs() : [];
            y = Array.isArray(ys) && Number.isFinite(ys[index]) ? ys[index] : start.y;
        }
        const yShift: number = Number(articulation.y_shift ?? 0);
        if (Number.isFinite(yShift)) {
            y += yShift;
        }

        let originY: number = position === positions.ABOVE ? 1 : 0;
        if (!isTab && (position === positions.ABOVE || position === positions.BELOW) && Array.isArray(note.getYs?.())) {
            const ys: number[] = note.getYs();
            const keyProps: any[] = typeof note.getKeyProps === "function" ? note.getKeyProps() : note.keyProps;
            const noteLine: number = Number(keyProps?.[index]?.line);
            if (Number.isFinite(ys[index]) && Number.isFinite(noteLine) && Number.isFinite(staffSpace) && staffSpace !== 0) {
                const offsetDirection: number = position === positions.ABOVE ? -1 : 1;
                const distanceFromNote: number = (ys[index] - y) / staffSpace;
                const articLine: number = distanceFromNote + noteLine;
                const snappedLine: number =
                    this.snapFingeringArticulationLineToStaff(canSitBetweenLines, articLine, position, offsetDirection);
                if (this.fingeringArticulationLineIsWithinStaff(snappedLine, position)) {
                    originY = 0.5;
                }
                y += Math.abs(snappedLine - articLine) * staffSpace * offsetDirection;
            }
        }

        const top: number = y - (height * originY);
        return this.normalizeFingeringPxRect({
            x: x - width / 2,
            y: top,
            width,
            height,
        }, 1.5);
    }

    private getFingeringOrnamentCollisionRectPx(ornament: any): FingeringPxRect {
        const note: any = typeof ornament?.getNote === "function" ? ornament.getNote() : ornament?.note;
        const index: number = Number(typeof ornament?.getIndex === "function" ? ornament.getIndex() : ornament?.index);
        const position: number = this.getFingeringModifierPosition(ornament);
        const stave: any = typeof note?.getStave === "function" ? note.getStave() : note?.stave;
        if (!note || !stave || !Number.isFinite(index) || !Number.isFinite(position)) {
            return undefined;
        }

        const start: any = typeof note.getModifierStartXY === "function"
            ? note.getModifierStartXY(position, index)
            : undefined;
        if (!start || !Number.isFinite(start.x)) {
            return undefined;
        }

        const stemDirection: number = this.getFingeringNoteStemDirection(note);
        const stemExtents: any = this.getFingeringNoteStemExtents(note);
        const spacing: number = Number(stave.getSpacingBetweenLines?.() ?? 10);
        const textLine: number = Number(ornament.text_line ?? 0);
        const isStemDown: boolean = stemDirection === this.getFingeringVexFlowStemDown();
        let baseY: number = isStemDown ? Number(stemExtents?.baseY) : Number(stemExtents?.topY);
        if (!Number.isFinite(baseY)) {
            const ys: number[] = typeof note.getYs === "function" ? note.getYs() : [];
            const finiteYs: number[] = Array.isArray(ys) ? ys.filter((yValue: number) => Number.isFinite(yValue)) : [];
            baseY = finiteYs.length > 0 ? Math.min(...finiteYs) : start.y;
        }

        if (note.getCategory?.() === "tabnotes") {
            if (typeof note.hasStem === "function" && note.hasStem()) {
                if (isStemDown) {
                    baseY = Number(stave.getYForTopText?.(textLine) ?? baseY);
                }
            } else {
                baseY = Number(stave.getYForTopText?.(textLine) ?? baseY);
            }
        }

        const isPlacedOnNoteheadSide: boolean = isStemDown;
        let lineSpacing: number = 1;
        if (!isPlacedOnNoteheadSide && note.beam) {
            lineSpacing += 0.5;
        }
        if (ornament?.ornament?.smuflGlyph && isPlacedOnNoteheadSide) {
            lineSpacing += 0.4;
        }

        let glyphX: number = start.x;
        if (ornament.delayed) {
            glyphX += this.getFingeringDelayedOrnamentXShiftPx(ornament, note, stave);
        }
        const glyphYBetweenLines: number = baseY - (spacing * (textLine + lineSpacing));
        let glyphY: number = Math.min(Number(stave.getYForTopText?.(textLine) ?? glyphYBetweenLines), glyphYBetweenLines);
        const yShift: number = Number(ornament.y_shift ?? 0);
        if (Number.isFinite(yShift)) {
            glyphY += yShift;
        }

        const ornamentMetrics: FingeringOrnamentVisualMetrics = this.getFingeringOrnamentGlyphMetrics(ornament);
        const lowerMetrics: { width: number, height: number } =
            this.getFingeringVexFlowGlyphMetrics(ornament.accidentalLower, undefined);
        const upperMetrics: { width: number, height: number } =
            this.getFingeringVexFlowGlyphMetrics(ornament.accidentalUpper, undefined);
        const lowerPadding: number = Math.max(Number(ornament.render_options?.accidentalLowerPadding ?? 3), 0);
        const upperPadding: number = Math.max(Number(ornament.render_options?.accidentalUpperPadding ?? 3), 0);
        const hasLower: boolean = lowerMetrics.width > 0.01 && lowerMetrics.height > 0.01;
        const hasUpper: boolean = upperMetrics.width > 0.01 && upperMetrics.height > 0.01;
        const width: number = Math.max(
            ornamentMetrics.width,
            hasLower ? lowerMetrics.width : 0,
            hasUpper ? upperMetrics.width : 0,
            Number(ornament?.getWidth?.() ?? ornament?.width ?? 0),
            5
        );
        const x: number = glyphX - width / 2;
        let cursorY: number = glyphY;
        let topY: number = Number.POSITIVE_INFINITY;
        let bottomY: number = Number.NEGATIVE_INFINITY;
        const includeYBounds: (top: number, bottom: number) => void = (top: number, bottom: number): void => {
            if (!Number.isFinite(top) || !Number.isFinite(bottom)) {
                return;
            }
            topY = Math.min(topY, top, bottom);
            bottomY = Math.max(bottomY, top, bottom);
        };

        if (hasLower) {
            includeYBounds(cursorY - lowerMetrics.height, cursorY);
            cursorY -= lowerMetrics.height + lowerPadding;
        }

        const ornamentBottom: number = cursorY + ornamentMetrics.bottomOffset;
        includeYBounds(ornamentBottom - Math.max(ornamentMetrics.height, 5), ornamentBottom);
        cursorY -= Math.max(ornamentMetrics.advanceHeight, ornamentMetrics.height, 5);

        if (hasUpper) {
            const upperY: number = cursorY - upperPadding;
            includeYBounds(upperY - upperMetrics.height, upperY);
        }

        if (!Number.isFinite(topY) || !Number.isFinite(bottomY) || bottomY <= topY) {
            return undefined;
        }

        return this.normalizeFingeringPxRect({
            x,
            y: topY,
            width,
            height: bottomY - topY,
        }, ornament?.ornament?.smuflGlyph ? 1 : 1.5);
    }

    private addFingeringBeamCollisionRects(rects: FingeringCollisionRect[], measure: GraphicalMeasure,
                                           line: StaffLine): void {
        const vexFlowMeasure: any = measure as any;
        const beams: any[] = [];
        const vfbeams: any = vexFlowMeasure.vfbeams;
        if (vfbeams) {
            for (const voiceID in vfbeams) {
                if (vfbeams.hasOwnProperty(voiceID) && Array.isArray(vfbeams[voiceID])) {
                    beams.push(...vfbeams[voiceID]);
                }
            }
        }
        if (Array.isArray(vexFlowMeasure.autoVfBeams)) {
            beams.push(...vexFlowMeasure.autoVfBeams);
        }
        if (Array.isArray(vexFlowMeasure.autoTupletVfBeams)) {
            beams.push(...vexFlowMeasure.autoTupletVfBeams);
        }

        const measureX: number = measure.PositionAndShape.RelativePosition.x;
        for (let beamIndex: number = 0; beamIndex < beams.length; beamIndex++) {
            const beam: any = beams[beamIndex];
            const beamRects: FingeringCollisionRect[] = this.getFingeringBeamCollisionRects(beam, measureX);
            for (let rectIndex: number = 0; rectIndex < beamRects.length; rectIndex++) {
                this.addRectIfValid(rects, this.withFingeringDebugSource(
                    beamRects[rectIndex],
                    measure,
                    line,
                    `beam:${beamIndex}:${rectIndex}`
                ));
            }
        }
    }

    private getFingeringBeamCollisionRects(beam: any, measureX: number): FingeringCollisionRect[] {
        if (!Array.isArray(beam?.notes) || beam.notes.length < 2) {
            return [];
        }
        this.ensureFingeringBeamFormatted(beam);
        if (typeof beam.getBeamYToDraw !== "function" || typeof beam.getSlopeY !== "function") {
            const fallbackRect: FingeringCollisionRect = this.getFingeringBeamCollisionRect(beam, measureX);
            return fallbackRect ? [fallbackRect] : [];
        }

        const firstNote: any = beam.notes[0];
        const firstStemX: number = Number(firstNote?.getStemX?.());
        if (!Number.isFinite(firstStemX)) {
            return [];
        }
        const staveOrigin: { x: number, y: number } = this.getVexFlowStaveOriginPx(firstNote);
        const stemWidth: number = Math.max(Number((VF.Stem as any)?.WIDTH ?? this.rules.StemWidth * 10), 0);
        const signedBeamThickness: number = Math.max(Number(beam.render_options?.beam_width ?? 5), 3) *
            (this.getFingeringBeamStemDirection(beam) < 0 ? -1 : 1);
        const validBeamDurations: string[] = ["4", "8", "16", "32", "64"];
        const rects: FingeringCollisionRect[] = [];
        let beamY: number = Number(beam.getBeamYToDraw());
        if (!Number.isFinite(beamY)) {
            return [];
        }

        for (const duration of validBeamDurations) {
            const beamLines: any[] = typeof beam.getBeamLines === "function" ? beam.getBeamLines(duration) : [];
            if (!Array.isArray(beamLines) || beamLines.length === 0) {
                beamY += signedBeamThickness * 1.5;
                continue;
            }
            for (const beamLine of beamLines) {
                const startX: number = Number(beamLine?.start);
                const endX: number = Number(beamLine?.end);
                if (!Number.isFinite(startX) || !Number.isFinite(endX)) {
                    continue;
                }
                const startY: number = Number(beam.getSlopeY(startX, firstStemX, beamY, beam.slope));
                const endY: number = Number(beam.getSlopeY(endX, firstStemX, beamY, beam.slope));
                if (!Number.isFinite(startY) || !Number.isFinite(endY)) {
                    continue;
                }
                const visualEndX: number = endX >= startX ? endX + stemWidth : endX - stemWidth;
                const segments: FingeringPxRect[] =
                    this.segmentFingeringBeamLineCollisionRectsPx(startX, startY, visualEndX, endY, signedBeamThickness);
                for (const segment of segments) {
                    const rect: FingeringCollisionRect =
                        this.fingeringPxRectToCollisionRect(segment, measureX, staveOrigin, "beam", 1);
                    if (rect && rect.bottom >= -4 && rect.top <= this.rules.StaffHeight + 4) {
                        rects.push(rect);
                    }
                }
            }
            beamY += signedBeamThickness * 1.5;
        }

        if (rects.length > 0) {
            return rects;
        }
        const fallbackBeamRect: FingeringCollisionRect = this.getFingeringBeamCollisionRect(beam, measureX);
        return fallbackBeamRect ? [fallbackBeamRect] : [];
    }

    private segmentFingeringBeamLineCollisionRectsPx(startX: number, startY: number, endX: number, endY: number,
                                                     signedThickness: number): FingeringPxRect[] {
        const dx: number = endX - startX;
        const segmentCount: number = Math.max(1, Math.min(64, Math.ceil(Math.abs(dx) / 8)));
        const rects: FingeringPxRect[] = [];
        for (let i: number = 0; i < segmentCount; i++) {
            const t0: number = i / segmentCount;
            const t1: number = (i + 1) / segmentCount;
            const x0: number = startX + dx * t0;
            const x1: number = startX + dx * t1;
            const y0: number = startY + (endY - startY) * t0;
            const y1: number = startY + (endY - startY) * t1;
            const rect: FingeringPxRect = this.normalizeFingeringPxRect({
                x: Math.min(x0, x1),
                y: Math.min(y0, y1, y0 + signedThickness, y1 + signedThickness),
                width: Math.abs(x1 - x0),
                height: Math.max(y0, y1, y0 + signedThickness, y1 + signedThickness) -
                    Math.min(y0, y1, y0 + signedThickness, y1 + signedThickness),
            }, 0.35);
            if (rect && rect.width > 0.01 && rect.height > 0.01) {
                rects.push(rect);
            }
        }
        return rects;
    }

    private getFingeringBeamCollisionRect(beam: any, measureX: number): FingeringCollisionRect {
        if (!Array.isArray(beam?.notes) || beam.notes.length < 2) {
            return undefined;
        }
        const stemXs: number[] = beam.notes.map((note: any) => note?.getStemX?.()).filter((x: number) => Number.isFinite(x));
        if (stemXs.length < 2 || typeof beam.getBeamYToDraw !== "function" || typeof beam.getSlopeY !== "function") {
            return undefined;
        }
        const minX: number = Math.min(...stemXs);
        const maxX: number = Math.max(...stemXs);
        const firstNote: any = beam.notes[0];
        if (!firstNote || typeof firstNote.getStemX !== "function") {
            return undefined;
        }
        const staveOrigin: { x: number, y: number } = this.getVexFlowStaveOriginPx(firstNote);
        const firstStemX: number = firstNote.getStemX();
        const beamY: number = beam.getBeamYToDraw();
        const beamThickness: number = Math.max(beam.render_options?.beam_width ?? 5, 3);
        const beamBandHeight: number = beamThickness * this.getFingeringBeamStemDirection(beam) *
            (((beam.beam_count ?? 1) - 1) * 1.5 + 1);
        const yAtMin: number = beam.getSlopeY(minX, firstStemX, beamY, beam.slope);
        const yAtMax: number = beam.getSlopeY(maxX, firstStemX, beamY, beam.slope);
        const top: number = (Math.min(yAtMin, yAtMax, yAtMin + beamBandHeight, yAtMax + beamBandHeight) - staveOrigin.y) / 10;
        const bottom: number = (Math.max(yAtMin, yAtMax, yAtMin + beamBandHeight, yAtMax + beamBandHeight) - staveOrigin.y) / 10;
        if (!Number.isFinite(top) || !Number.isFinite(bottom) || bottom < -4 || top > this.rules.StaffHeight + 4) {
            return undefined;
        }
        return {
            left: measureX + (minX - staveOrigin.x) / 10,
            right: measureX + (maxX - staveOrigin.x) / 10 + 0.12,
            top,
            bottom,
            penalty: 1,
            kind: "beam"
        };
    }

    private ensureFingeringBeamFormatted(beam: any): void {
        if (!beam || beam.postFormatted || typeof beam.postFormat !== "function") {
            return;
        }
        try {
            beam.postFormat();
        } catch (e) {
            // VexFlow can reject incomplete beams before draw; use any geometry already present.
        }
    }

    private getFingeringBeamStemDirection(beam: any): number {
        const direction: number = beam?.stem_direction ?? beam?.notes?.[0]?.getStemDirection?.();
        return direction < 0 ? -1 : 1;
    }

    private getFingeringVexFlowTieDirection(vfTie: any, graphicalTie: GraphicalTie): number {
        const direction: number = Number(vfTie?.direction);
        if (direction === 1 || direction === -1) {
            return direction < 0 ? -1 : 1;
        }
        const firstDirection: number = Number(vfTie?.first_note?.getStemDirection?.());
        if (firstDirection === 1 || firstDirection === -1) {
            return firstDirection;
        }
        const lastDirection: number = Number(vfTie?.last_note?.getStemDirection?.());
        if (lastDirection === 1 || lastDirection === -1) {
            return lastDirection;
        }
        const tieDirection: PlacementEnum = graphicalTie?.Tie?.getTieDirection(graphicalTie?.StartNote?.sourceNote);
        if (tieDirection !== undefined) {
            return tieDirection === PlacementEnum.Below ? 1 : -1;
        }
        return 1;
    }

    private getFingeringVexFlowTieGeometry(vfTie: any): {
        firstX: number;
        lastX: number;
        firstYs: number[];
        lastYs: number[];
        firstIndices: number[];
        lastIndices: number[];
        firstNote: any;
        lastNote: any;
    } {
        const firstNote: any = vfTie?.first_note;
        const lastNote: any = vfTie?.last_note;
        let firstX: number;
        let lastX: number;
        let firstYs: number[];
        let lastYs: number[];
        let firstIndices: number[] = Array.isArray(vfTie?.first_indices) ? vfTie.first_indices : [0];
        let lastIndices: number[] = Array.isArray(vfTie?.last_indices) ? vfTie.last_indices : [0];

        if (firstNote) {
            firstX = firstNote.getTieRightX() + (vfTie?.render_options?.tie_spacing ?? 0);
            firstYs = firstNote.getYs();
        } else if (lastNote) {
            firstX = lastNote.getStave().getTieStartX();
            firstYs = lastNote.getYs();
            firstIndices = lastIndices;
        } else {
            return undefined;
        }

        if (lastNote) {
            lastX = lastNote.getTieLeftX() + (vfTie?.render_options?.tie_spacing ?? 0);
            lastYs = lastNote.getYs();
        } else if (firstNote) {
            lastX = firstNote.getStave().getTieEndX();
            lastYs = firstNote.getYs();
            lastIndices = firstIndices;
        } else {
            return undefined;
        }

        if (!Number.isFinite(firstX) || !Number.isFinite(lastX) ||
            !Array.isArray(firstYs) || !Array.isArray(lastYs) ||
            firstYs.length === 0 || lastYs.length === 0) {
                return undefined;
        }
        return { firstX, lastX, firstYs, lastYs, firstIndices, lastIndices, firstNote, lastNote };
    }

    private getFingeringTieReferenceGraphicalNote(graphicalTie: GraphicalTie,
                                                  geometry: ReturnType<typeof this.getFingeringVexFlowTieGeometry>): GraphicalNote {
        if (!geometry) {
            return undefined;
        }
        if (geometry.firstNote && graphicalTie?.StartNote) {
            return graphicalTie.StartNote;
        }
        if (geometry.lastNote && graphicalTie?.EndNote) {
            return graphicalTie.EndNote;
        }
        return graphicalTie?.StartNote ?? graphicalTie?.EndNote;
    }

    private getFingeringVexFlowTieCollisionRect(vfTie: any, line: StaffLine,
                                                referenceMeasure: GraphicalMeasure,
                                                graphicalTie?: GraphicalTie): FingeringCollisionRect {
        const geometry: ReturnType<typeof this.getFingeringVexFlowTieGeometry> =
            this.getFingeringVexFlowTieGeometry(vfTie);
        if (!geometry) {
            return undefined;
        }

        let collisionMeasure: GraphicalMeasure = referenceMeasure;
        if (graphicalTie) {
            const referenceGraphicalNote: GraphicalNote =
                this.getFingeringTieReferenceGraphicalNote(graphicalTie, geometry);
            collisionMeasure = referenceGraphicalNote?.parentVoiceEntry?.parentStaffEntry?.parentMeasure;
            if (!referenceGraphicalNote) {
                return undefined;
            }
        }
        if (!collisionMeasure || collisionMeasure.ParentStaffLine !== line) {
            return undefined;
        }

        const direction: number = this.getFingeringVexFlowTieDirection(vfTie, graphicalTie);
        const yShift: number = Number.isFinite(vfTie?.render_options?.y_shift) ? vfTie.render_options.y_shift : 7;
        const cp2: number = Number.isFinite(vfTie?.render_options?.cp2) ? vfTie.render_options.cp2 : 12;
        const yValues: number[] = [];
        const indexCount: number = Math.max(geometry.firstIndices.length, geometry.lastIndices.length);
        for (let tieNoteIndex: number = 0; tieNoteIndex < indexCount; tieNoteIndex++) {
            const firstIndex: number = geometry.firstIndices[Math.min(tieNoteIndex, geometry.firstIndices.length - 1)];
            const lastIndex: number = geometry.lastIndices[Math.min(tieNoteIndex, geometry.lastIndices.length - 1)];
            const firstY: number = geometry.firstYs[firstIndex] + yShift * direction;
            const lastY: number = geometry.lastYs[lastIndex] + yShift * direction;
            const controlY: number = (firstY + lastY) / 2 + cp2 * direction;
            if (Number.isFinite(firstY) && Number.isFinite(lastY) && Number.isFinite(controlY)) {
                yValues.push(firstY, lastY, controlY);
            }
        }
        if (yValues.length === 0) {
            return undefined;
        }

        const staveOrigin: { x: number, y: number } =
            this.getVexFlowStaveOriginPx(geometry.firstNote ?? geometry.lastNote);
        return this.fingeringPxRectToCollisionRect({
            x: Math.min(geometry.firstX, geometry.lastX),
            y: Math.min(...yValues) - 3,
            width: Math.abs(geometry.lastX - geometry.firstX),
            height: Math.max(...yValues) - Math.min(...yValues) + 6
        }, collisionMeasure.PositionAndShape.RelativePosition.x, staveOrigin, "tie", 1.35);
    }

    private getFingeringTieVexFlowCollisionRect(graphicalTie: GraphicalTie, line: StaffLine): FingeringCollisionRect {
        return this.getFingeringVexFlowTieCollisionRect(graphicalTie?.vfTie as any, line, undefined, graphicalTie);
    }

    private getFingeringTieFallbackCollisionRect(graphicalTie: GraphicalTie, line: StaffLine): FingeringCollisionRect {
        const startNote: GraphicalNote = graphicalTie.StartNote;
        const endNote: GraphicalNote = graphicalTie.EndNote;
        const startRect: FingeringCollisionRect = startNote ?
            this.getBoundingBoxRectInStaffLine(startNote.PositionAndShape, line, true) : undefined;
        const endRect: FingeringCollisionRect = endNote ?
            this.getBoundingBoxRectInStaffLine(endNote.PositionAndShape, line, true) : undefined;
        if (!startRect && !endRect) {
            return undefined;
        }
        const tieDirection: PlacementEnum = graphicalTie.Tie.getTieDirection(startNote?.sourceNote);
        const left: number = Math.min(startRect?.left ?? endRect.left, endRect?.left ?? startRect.left) - 0.22;
        const right: number = Math.max(startRect?.right ?? endRect.right, endRect?.right ?? startRect.right) + 0.22;
        const referenceTop: number = Math.min(startRect?.top ?? endRect.top, endRect?.top ?? startRect.top);
        const referenceBottom: number = Math.max(startRect?.bottom ?? endRect.bottom, endRect?.bottom ?? startRect.bottom);
        const centerY: number = tieDirection === PlacementEnum.Below ? referenceBottom + 0.48 : referenceTop - 0.48;
        const tieThickness: number = 0.34;
        const tieArchHeight: number = 2.05;
        return {
            left,
            right,
            top: tieDirection === PlacementEnum.Below ? centerY - tieThickness : centerY - tieArchHeight,
            bottom: tieDirection === PlacementEnum.Below ? centerY + tieArchHeight : centerY + tieThickness,
            penalty: 1.15,
            kind: "tie"
        };
    }

    private getFingeringGraphicalNoteVexFlowNote(note: GraphicalNote): any {
        return (note as any)?.vfnote?.[0];
    }

    private getFingeringGraphicalNoteVexFlowIndex(note: GraphicalNote): number {
        const noteIndex: number = Number((note as any)?.vfnote?.[1] ?? (note as any)?.vfnoteIndex ?? 0);
        return Number.isFinite(noteIndex) ? noteIndex : 0;
    }

    private addFingeringSplitTieSegmentCollisionRects(rects: FingeringCollisionRect[], graphicalTie: GraphicalTie,
                                                      line: StaffLine, tieIndex: number): void {
        const startMeasure: GraphicalMeasure =
            graphicalTie.StartNote?.parentVoiceEntry?.parentStaffEntry?.parentMeasure;
        const endMeasure: GraphicalMeasure =
            graphicalTie.EndNote?.parentVoiceEntry?.parentStaffEntry?.parentMeasure;
        if (!startMeasure || !endMeasure || startMeasure.ParentStaffLine === endMeasure.ParentStaffLine) {
            return;
        }

        const addSegmentRect: (vfTie: any, debugMeasure: GraphicalMeasure, debugKey: string) => void =
            (vfTie: any, debugMeasure: GraphicalMeasure, debugKey: string): void => {
                const tieRect: FingeringCollisionRect =
                    this.getFingeringVexFlowTieCollisionRect(vfTie, line, undefined, graphicalTie);
                this.addRectIfValid(rects, this.withFingeringDebugSource(tieRect, debugMeasure, line, debugKey));
            };

        if (startMeasure.ParentStaffLine === line) {
            const startVfNote: any = this.getFingeringGraphicalNoteVexFlowNote(graphicalTie.StartNote);
            if (startVfNote) {
                addSegmentRect({
                    first_note: startVfNote,
                    first_indices: [this.getFingeringGraphicalNoteVexFlowIndex(graphicalTie.StartNote)],
                    direction: (graphicalTie.vfTie as any)?.direction,
                    render_options: (graphicalTie.vfTie as any)?.render_options
                }, startMeasure, `tie-start:${tieIndex}`);
            }
        }

        if (endMeasure.ParentStaffLine === line) {
            const endVfNote: any = this.getFingeringGraphicalNoteVexFlowNote(graphicalTie.EndNote);
            if (endVfNote) {
                addSegmentRect({
                    last_note: endVfNote,
                    last_indices: [this.getFingeringGraphicalNoteVexFlowIndex(graphicalTie.EndNote)],
                    direction: (graphicalTie.vfTie as any)?.direction,
                    render_options: (graphicalTie.vfTie as any)?.render_options
                }, endMeasure, `tie-end:${tieIndex}`);
            }
        }
    }

    private addFingeringTieCollisionRects(rects: FingeringCollisionRect[], gse: GraphicalStaffEntry,
                                          line: StaffLine, measure: GraphicalMeasure): void {
        for (let tieIndex: number = 0; tieIndex < gse.GraphicalTies.length; tieIndex++) {
            const graphicalTie: GraphicalTie = gse.GraphicalTies[tieIndex];
            const vexFlowTieRect: FingeringCollisionRect =
                this.getFingeringTieVexFlowCollisionRect(graphicalTie, line);
            this.addRectIfValid(rects, this.withFingeringDebugSource(vexFlowTieRect, measure, line, `tie:${tieIndex}`));
            const fallbackTieRect: FingeringCollisionRect =
                this.getFingeringTieFallbackCollisionRect(graphicalTie, line);
            this.addRectIfValid(rects, this.withFingeringDebugSource(fallbackTieRect, measure, line, `tie-fallback:${tieIndex}`));
            this.addFingeringSplitTieSegmentCollisionRects(rects, graphicalTie, line, tieIndex);
        }
    }

    private addFingeringCrossLineTieCollisionRects(rects: FingeringCollisionRect[], line: StaffLine): void {
        const seenTies: Set<GraphicalTie> = new Set<GraphicalTie>();
        let tieIndex: number = 0;
        for (const measureColumn of this.graphicalMusicSheet?.MeasureList ?? []) {
            for (const measure of measureColumn ?? []) {
                for (const staffEntry of measure?.staffEntries ?? []) {
                    for (const graphicalTie of staffEntry.GraphicalTies ?? []) {
                        if (seenTies.has(graphicalTie)) {
                            continue;
                        }
                        seenTies.add(graphicalTie);
                        const startLine: StaffLine =
                            graphicalTie.StartNote?.parentVoiceEntry?.parentStaffEntry?.parentMeasure?.ParentStaffLine;
                        const endLine: StaffLine =
                            graphicalTie.EndNote?.parentVoiceEntry?.parentStaffEntry?.parentMeasure?.ParentStaffLine;
                        if (startLine && endLine && startLine !== endLine && (startLine === line || endLine === line)) {
                            this.addFingeringSplitTieSegmentCollisionRects(rects, graphicalTie, line, tieIndex);
                            tieIndex++;
                        }
                    }
                }
            }
        }
    }

    private addFingeringVexFlowMeasureTieCollisionRects(rects: FingeringCollisionRect[], measure: GraphicalMeasure,
                                                        line: StaffLine): boolean {
        const vfTies: any[] = (measure as any)?.vfTies;
        if (!Array.isArray(vfTies) || vfTies.length === 0) {
            return false;
        }

        let addedTieRect: boolean = false;
        for (let tieIndex: number = 0; tieIndex < vfTies.length; tieIndex++) {
            const tieRect: FingeringCollisionRect =
                this.getFingeringVexFlowTieCollisionRect(vfTies[tieIndex], line, measure);
            const previousLength: number = rects.length;
            this.addRectIfValid(rects, this.withFingeringDebugSource(tieRect, measure, line, `vf-tie:${tieIndex}`));
            addedTieRect = addedTieRect || rects.length > previousLength;
        }
        return addedTieRect;
    }

    private getFingeringOrnamentReserveSize(ornamentContainer: OrnamentContainer): { halfWidth: number, height: number } {
        const ornamentName: string = this.getFingeringVexflowOrnamentName(ornamentContainer);
        const hasAccidental: boolean =
            ornamentContainer.AccidentalAbove !== AccidentalEnum.NONE ||
            ornamentContainer.AccidentalBelow !== AccidentalEnum.NONE;
        const wideOrnament: boolean = /turn|prall|mordent|tremblement/.test(ornamentName ?? "");
        const halfWidth: number = wideOrnament ? 1.35 : 1.05;
        const height: number = (wideOrnament ? 1.8 : 1.45) + (hasAccidental ? 0.85 : 0);
        return { halfWidth, height };
    }

    private getFingeringStemLaneExtremeY(vfNote: any, above: boolean): number {
        const stemExtents: any = this.getFingeringNoteStemExtents(vfNote);
        const ys: number[] = [Number(stemExtents?.topY), Number(stemExtents?.baseY)];
        const noteYs: number[] = typeof vfNote?.getYs === "function" ? vfNote.getYs() : [];
        if (Array.isArray(noteYs)) {
            ys.push(...noteYs.map((y: number) => Number(y)));
        }
        const finiteYs: number[] = ys.filter((y: number) => Number.isFinite(y));
        if (finiteYs.length === 0) {
            return undefined;
        }
        const staveOrigin: { x: number, y: number } = this.getVexFlowStaveOriginPx(vfNote);
        const extremeY: number = above ? Math.min(...finiteYs) : Math.max(...finiteYs);
        return (extremeY - staveOrigin.y) / 10;
    }

    private adjustFingeringOrnamentReserveToStemLane(rect: FingeringCollisionRect,
                                                     ornamentContainer: OrnamentContainer,
                                                     vfNote: any): FingeringCollisionRect {
        if (!rect) {
            return rect;
        }
        const above: boolean = ornamentContainer.placement !== PlacementEnum.Below;
        const stemLaneY: number = this.getFingeringStemLaneExtremeY(vfNote, above);
        if (!Number.isFinite(stemLaneY)) {
            return rect;
        }
        const clearance: number = 0.18;
        const targetEdge: number = above ? stemLaneY - clearance : stemLaneY + clearance;
        if ((above && rect.bottom <= targetEdge + 0.03) ||
            (!above && rect.top >= targetEdge - 0.03)) {
            return rect;
        }

        const size: { halfWidth: number, height: number } = this.getFingeringOrnamentReserveSize(ornamentContainer);
        const centerX: number = (rect.left + rect.right) / 2;
        const halfWidth: number = Math.max((rect.right - rect.left) / 2, size.halfWidth);
        const height: number = Math.max(rect.bottom - rect.top, size.height);
        return {
            left: centerX - halfWidth,
            right: centerX + halfWidth,
            top: above ? targetEdge - height : targetEdge,
            bottom: above ? targetEdge : targetEdge + height,
            penalty: rect.penalty,
            kind: rect.kind
        };
    }

    private getFingeringOrnamentReserveRect(voiceEntry: GraphicalVoiceEntry, line: StaffLine,
                                            measure: GraphicalMeasure): FingeringCollisionRect {
        const ornamentContainer: OrnamentContainer = voiceEntry.parentVoiceEntry?.OrnamentContainer;
        const graphicalNote: GraphicalNote = voiceEntry.notes[0];
        if (!ornamentContainer || !graphicalNote) {
            return undefined;
        }

        const vfNote: any = (graphicalNote as any).vfnote?.[0] ?? (voiceEntry as any).vfStaveNote;
        if (vfNote) {
            try {
                const estimate: any = this.createFingeringOrnamentEstimateModifier(ornamentContainer, vfNote);
                const pxRect: FingeringPxRect = this.getFingeringOrnamentCollisionRectPx(estimate);
                const staveOrigin: { x: number, y: number } = this.getVexFlowStaveOriginPx(vfNote);
                const measureX: number = measure.PositionAndShape.RelativePosition.x;
                const rect: FingeringCollisionRect =
                    this.fingeringPxRectToCollisionRect(pxRect, measureX, staveOrigin, "ornament-reserve", 1);
                if (rect) {
                    return this.adjustFingeringOrnamentReserveToStemLane(rect, ornamentContainer, vfNote);
                }
            } catch (e) {
                // Fall through to the staff-relative estimate when VexFlow geometry is incomplete.
            }
        }

        const noteRect: FingeringCollisionRect =
            this.getBoundingBoxRectInStaffLine(graphicalNote.PositionAndShape, line, true);
        if (!noteRect) {
            return undefined;
        }
        const centerX: number = (noteRect.left + noteRect.right) / 2;
        const above: boolean = ornamentContainer.placement !== PlacementEnum.Below;
        const size: { halfWidth: number, height: number } = this.getFingeringOrnamentReserveSize(ornamentContainer);
        const stemLaneOffset: number = 3.1;
        return {
            left: centerX - size.halfWidth,
            right: centerX + size.halfWidth,
            top: above ? noteRect.top - stemLaneOffset - size.height : noteRect.bottom + stemLaneOffset,
            bottom: above ? noteRect.top - stemLaneOffset : noteRect.bottom + stemLaneOffset + size.height,
            penalty: 1,
            kind: "ornament-reserve"
        };
    }

    private addFingeringOrnamentReserveRects(rects: FingeringCollisionRect[], gse: GraphicalStaffEntry,
                                             line: StaffLine, measure: GraphicalMeasure): void {
        for (let voiceEntryIndex: number = 0; voiceEntryIndex < gse.graphicalVoiceEntries.length; voiceEntryIndex++) {
            const voiceEntry: GraphicalVoiceEntry = gse.graphicalVoiceEntries[voiceEntryIndex];
            if (!voiceEntry.parentVoiceEntry?.OrnamentContainer || voiceEntry.notes.length === 0) {
                continue;
            }
            this.addRectIfValid(rects, this.withFingeringDebugSource(
                this.getFingeringOrnamentReserveRect(voiceEntry, line, measure),
                measure,
                line,
                `ornament-reserve:${voiceEntryIndex}`
            ));
        }
    }

    private getFingeringCollisionRects(system: MusicSystem, line: StaffLine, measure: GraphicalMeasure,
                                       placement: PlacementEnum): FingeringCollisionRect[] {
        const rects: FingeringCollisionRect[] = [];
        const currentMeasureIndex: number = line.Measures.indexOf(measure);
        const firstMeasureIndex: number = Math.max(0, currentMeasureIndex - 1);
        const lastMeasureIndex: number = Math.min(line.Measures.length - 1, currentMeasureIndex + 1);
        for (let measureIndex: number = firstMeasureIndex; measureIndex <= lastMeasureIndex; measureIndex++) {
            const nearbyMeasure: GraphicalMeasure = line.Measures[measureIndex];
            this.addFingeringBeamCollisionRects(rects, nearbyMeasure, line);
            const hasMeasureTieRects: boolean =
                this.addFingeringVexFlowMeasureTieCollisionRects(rects, nearbyMeasure, line);
            for (const staffEntry of nearbyMeasure.staffEntries) {
                for (const voiceEntry of staffEntry.graphicalVoiceEntries) {
                    for (const note of voiceEntry.notes) {
                        if (note.sourceNote?.isRest()) {
                            this.addFingeringRestCollisionRects(rects, note, line, nearbyMeasure);
                        } else {
                            this.addFingeringNoteCollisionRects(rects, note, line, nearbyMeasure);
                        }
                    }
                }
                for (let fingeringIndex: number = 0; fingeringIndex < staffEntry.FingeringEntries.length; fingeringIndex++) {
                    const existingFingering: GraphicalLabel = staffEntry.FingeringEntries[fingeringIndex];
                    this.addRectIfValid(rects, this.withFingeringDebugSource(
                        this.getExistingFingeringCollisionRect(existingFingering, placement),
                        nearbyMeasure,
                        line,
                        `existing-fingering:${measureIndex}:${fingeringIndex}`
                    ));
                }
                const attachedModifierKinds: Set<string> =
                    this.addFingeringAttachedModifierCollisionRects(rects, staffEntry, line, nearbyMeasure);
                if (!attachedModifierKinds.has("ornament")) {
                    this.addFingeringOrnamentReserveRects(rects, staffEntry, line, nearbyMeasure);
                }
                if (!hasMeasureTieRects) {
                    this.addFingeringTieCollisionRects(rects, staffEntry, line, nearbyMeasure);
                }
            }
        }
        this.addFingeringCrossLineTieCollisionRects(rects, line);
        for (let labelIndex: number = 0; labelIndex < system.MeasureNumberLabels.length; labelIndex++) {
            const measureNumberLabel: GraphicalLabel = system.MeasureNumberLabels[labelIndex];
            this.addRectIfValid(rects, this.withFingeringDebugSource(
                this.getMeasureNumberRectInStaffLine(measureNumberLabel, line),
                measure,
                line,
                `measure-number:${labelIndex}`
            ));
        }
        return rects;
    }

    private getFingeringBlockingCollisionKinds(rect: FingeringCollisionRect,
                                               collisionRects: FingeringCollisionRect[]): Set<string> {
        const padding: number = 0.08;
        const labelRect: FingeringCollisionRect = this.expandFingeringRect(rect, padding);
        const kinds: Set<string> = new Set<string>();
        for (const collisionRect of collisionRects) {
            const obstacleRect: FingeringCollisionRect = this.expandFingeringRect(collisionRect, padding);
            if (this.getRectOverlapArea(labelRect, obstacleRect) > 0) {
                kinds.add(collisionRect.kind ?? "unknown");
            }
        }
        return kinds;
    }

    private getRelevantFingeringGroupCollisionRects(items: FingeringPlacementWorkItem[],
                                                    collisionRects: FingeringCollisionRect[],
                                                    horizontalSearchRadius: number): FingeringCollisionRect[] {
        const padding: number = 0.12;
        const searchLeft: number = Math.min(...items.map((item: FingeringPlacementWorkItem) => item.initialRect.left)) -
            horizontalSearchRadius - padding;
        const searchRight: number = Math.max(...items.map((item: FingeringPlacementWorkItem) => item.initialRect.right)) +
            horizontalSearchRadius + padding;
        return collisionRects.filter((rect: FingeringCollisionRect) => rect.right >= searchLeft && rect.left <= searchRight);
    }

    private getFingeringHorizontalNudgeOffsets(baseRect: FingeringCollisionRect, stacked: boolean): number[] {
        const labelWidth: number = Math.max(0.1, baseRect.right - baseRect.left);
        const step: number = Math.max(0.12, Math.min(stacked ? 0.18 : 0.22, labelWidth * 0.2));
        const maxShift: number = stacked ?
            Math.max(0.22, Math.min(0.36, labelWidth * 0.36)) :
            Math.max(0.32, Math.min(0.58, labelWidth * 0.5));
        const offsets: number[] = [0];
        for (let shift: number = step; shift <= maxShift + 0.0001; shift += step) {
            offsets.push(-shift, shift);
        }
        return offsets;
    }

    private shouldTryFingeringHorizontalNudge(blockingKinds: Set<string>, allowNoteBody: boolean): boolean {
        if (blockingKinds.size === 0) {
            return false;
        }
        for (const kind of blockingKinds) {
            if (kind !== "stem" && kind !== "notehead" && !(allowNoteBody && kind === "note")) {
                return false;
            }
        }
        return true;
    }

    private getFingeringVerticalSearchOffsets(items: FingeringPlacementWorkItem[], placement: PlacementEnum): number[] {
        const outwardSign: number = placement === PlacementEnum.Above ? -1 : 1;
        const maxLabelHeight: number = Math.max(...items.map((item: FingeringPlacementWorkItem) => item.labelHeight));
        const maxStackIndex: number = Math.max(...items.map((item: FingeringPlacementWorkItem) => item.stackIndex));
        const maxOutwardShift: number = Math.max(
            3.5,
            Math.min(6.2, maxLabelHeight * 1.1 + maxStackIndex * this.getFingeringStackSpacing(maxLabelHeight) + 3.2)
        );
        const offsets: number[] = [0];
        for (let outwardShift: number = 0.04; outwardShift <= maxOutwardShift + 0.0001;) {
            offsets.push(outwardShift * outwardSign);
            outwardShift += outwardShift < 2.4 ? 0.04 : 0.12;
        }
        return offsets;
    }

    private isFingeringOrnamentCollisionKind(kind: string): boolean {
        return kind === "ornament" || kind === "ornament-reserve";
    }

    private isFingeringHardCollisionKind(kind: string): boolean {
        return kind === "beam" || kind === "notehead" || kind === "stem" || kind === "tie";
    }

    private isFingeringTargetedVerticalCollisionKind(kind: string): boolean {
        return this.isFingeringOrnamentCollisionKind(kind) || this.isFingeringHardCollisionKind(kind);
    }

    private getFingeringCollisionClearancePadding(kind: string): number {
        switch (kind) {
            case "ornament":
            case "ornament-reserve":
                return 0.18;
            case "tie":
                return 0.18;
            case "beam":
                return 0.12;
            case "notehead":
                return 0.10;
            case "stem":
                return 0.08;
            default:
                return 0.035;
        }
    }

    private getFingeringCollisionClearanceFactor(kind: string): number {
        if (this.isFingeringOrnamentCollisionKind(kind)) {
            return 95;
        }
        if (kind === "tie") {
            return 60;
        }
        if (this.isFingeringHardCollisionKind(kind)) {
            return 44;
        }
        return 16;
    }

    private getFingeringGroupEnvelopeRect(items: FingeringPlacementWorkItem[]): FingeringCollisionRect {
        return {
            left: Math.min(...items.map((item: FingeringPlacementWorkItem) => item.initialRect.left)),
            right: Math.max(...items.map((item: FingeringPlacementWorkItem) => item.initialRect.right)),
            top: Math.min(...items.map((item: FingeringPlacementWorkItem) => item.initialRect.top)),
            bottom: Math.max(...items.map((item: FingeringPlacementWorkItem) => item.initialRect.bottom)),
            penalty: 1
        };
    }

    private getFingeringCandidateEnvelopeRect(candidates: FingeringLabelCandidate[]): FingeringCollisionRect {
        return {
            left: Math.min(...candidates.map((candidate: FingeringLabelCandidate) => candidate.rect.left)),
            right: Math.max(...candidates.map((candidate: FingeringLabelCandidate) => candidate.rect.right)),
            top: Math.min(...candidates.map((candidate: FingeringLabelCandidate) => candidate.rect.top)),
            bottom: Math.max(...candidates.map((candidate: FingeringLabelCandidate) => candidate.rect.bottom)),
            penalty: 1
        };
    }

    private getFingeringPlacedGroupEnvelopeRect(items: FingeringPlacementWorkItem[]): FingeringCollisionRect {
        const placedRects: FingeringCollisionRect[] = items
            .map((item: FingeringPlacementWorkItem) => item.placedRect)
            .filter((rect: FingeringCollisionRect) => !!rect);
        if (placedRects.length === 0) {
            return undefined;
        }
        return {
            left: Math.min(...placedRects.map((rect: FingeringCollisionRect) => rect.left)),
            right: Math.max(...placedRects.map((rect: FingeringCollisionRect) => rect.right)),
            top: Math.min(...placedRects.map((rect: FingeringCollisionRect) => rect.top)),
            bottom: Math.max(...placedRects.map((rect: FingeringCollisionRect) => rect.bottom)),
            penalty: 1
        };
    }

    private addFingeringTargetedVerticalOffsets(offsets: number[], rect: FingeringCollisionRect,
                                                placement: PlacementEnum,
                                                collisionRects: FingeringCollisionRect[]): void {
        const maxTargetedShift: number = 9.2;
        const maxOppositeSideShift: number = 6.2;
        const outwardSign: number = placement === PlacementEnum.Above ? -1 : 1;
        for (const collisionRect of collisionRects) {
            const kind: string = collisionRect.kind ?? "unknown";
            if (!this.isFingeringTargetedVerticalCollisionKind(kind)) {
                continue;
            }
            const padding: number = this.getFingeringCollisionClearancePadding(kind);
            const horizontalOverlap: number =
                Math.min(rect.right, collisionRect.right) + padding -
                (Math.max(rect.left, collisionRect.left) - padding);
            if (horizontalOverlap <= 0) {
                continue;
            }
            const outwardTargetDy: number = placement === PlacementEnum.Above
                ? collisionRect.top - rect.bottom - padding * 2
                : collisionRect.bottom - rect.top + padding * 2;
            if (!((placement === PlacementEnum.Above && outwardTargetDy >= -0.0001) ||
                (placement === PlacementEnum.Below && outwardTargetDy <= 0.0001) ||
                Math.abs(outwardTargetDy) > maxTargetedShift)) {
                    offsets.push(
                        outwardTargetDy,
                        outwardTargetDy + outwardSign * 0.04,
                        outwardTargetDy + outwardSign * 0.12
                    );
            }

            if (!this.isFingeringHardCollisionKind(kind)) {
                continue;
            }
            const oppositeTargetDy: number = placement === PlacementEnum.Above
                ? collisionRect.bottom - rect.top + padding * 2
                : collisionRect.top - rect.bottom - padding * 2;
            if ((placement === PlacementEnum.Above && oppositeTargetDy <= 0.0001) ||
                (placement === PlacementEnum.Below && oppositeTargetDy >= -0.0001) ||
                Math.abs(oppositeTargetDy) > maxOppositeSideShift) {
                    continue;
            }
            offsets.push(
                oppositeTargetDy,
                oppositeTargetDy - outwardSign * 0.04,
                oppositeTargetDy - outwardSign * 0.12
            );
        }
    }

    private getFingeringVerticalSearchOffsetsForCollisions(items: FingeringPlacementWorkItem[],
                                                           placement: PlacementEnum,
                                                           collisionRects: FingeringCollisionRect[]): number[] {
        const offsets: number[] = this.getFingeringVerticalSearchOffsets(items, placement);
        for (const item of items) {
            this.addFingeringTargetedVerticalOffsets(offsets, item.initialRect, placement, collisionRects);
        }
        if (items.length > 1) {
            this.addFingeringTargetedVerticalOffsets(
                offsets,
                this.getFingeringGroupEnvelopeRect(items),
                placement,
                collisionRects
            );
        }
        const uniqueOffsets: number[] = [];
        const seen: Set<string> = new Set<string>();
        for (const offset of offsets) {
            const key: string = offset.toFixed(3);
            if (seen.has(key)) {
                continue;
            }
            seen.add(key);
            uniqueOffsets.push(offset);
        }
        return uniqueOffsets.sort((left: number, right: number) =>
            Math.abs(left) - Math.abs(right) || left - right);
    }

    private getFingeringCollisionKindWeight(kind: string, stacked: boolean): number {
        switch (kind) {
            case "existing-fingering":
                return 90;
            case "measure-number":
                return 70;
            case "notehead":
                return stacked ? 95 : 46;
            case "rest":
                return 30;
            case "stem":
                return stacked ? 70 : 28;
            case "note":
                return 16;
            case "beam":
                return stacked ? 85 : 18;
            case "ornament":
                return 130;
            case "articulation":
                return 18;
            case "ornament-reserve":
                return 115;
            case "tie":
                return stacked ? 90 : 60;
            default:
                return 14;
        }
    }

    private getFingeringCollisionCost(rect: FingeringCollisionRect,
                                      collisionRects: FingeringCollisionRect[],
                                      stacked: boolean): number {
        let cost: number = 0;
        for (const collisionRect of collisionRects) {
            const kind: string = collisionRect.kind ?? "unknown";
            const padding: number = this.getFingeringCollisionClearancePadding(kind);
            const labelRect: FingeringCollisionRect = this.expandFingeringRect(rect, padding);
            const obstacleRect: FingeringCollisionRect = this.expandFingeringRect(collisionRect, padding);
            const paddedOverlap: number = this.getRectOverlapArea(labelRect, obstacleRect);
            if (paddedOverlap <= 0) {
                continue;
            }
            const actualOverlap: number = this.getRectOverlapArea(rect, collisionRect);
            const kindWeight: number = this.getFingeringCollisionKindWeight(kind, stacked);
            const rectPenalty: number = collisionRect.penalty ?? 1;
            const softFactor: number = collisionRect.soft ? 0.45 : 1;
            const clearanceFactor: number = this.getFingeringCollisionClearanceFactor(kind);
            cost += (actualOverlap * 120 + paddedOverlap * clearanceFactor) * kindWeight * rectPenalty * softFactor;
        }
        return cost;
    }

    private getFingeringVerticalMovementCost(dy: number, groupSize: number): number {
        const magnitude: number = Math.abs(dy);
        const stackMultiplier: number = 1 + Math.max(0, groupSize - 1) * 0.68;
        const softStackLimit: number = groupSize > 1 ? 2.15 : Number.POSITIVE_INFINITY;
        const excessStackShift: number = Math.max(0, magnitude - softStackLimit);
        return (magnitude * 9 + magnitude * magnitude * 3.2) * stackMultiplier +
            excessStackShift * 60 + excessStackShift * excessStackShift * 220;
    }

    private getFingeringHorizontalMovementCost(dx: number, stacked: boolean): number {
        const magnitude: number = Math.abs(dx);
        const linearWeight: number = stacked ? 85 : 40;
        const quadraticWeight: number = stacked ? 140 : 70;
        return magnitude * linearWeight + magnitude * magnitude * quadraticWeight;
    }

    private getFingeringCandidateHorizontalOffsets(rect: FingeringCollisionRect,
                                                   collisionRects: FingeringCollisionRect[],
                                                   stacked: boolean): number[] {
        const blockingKinds: Set<string> = this.getFingeringBlockingCollisionKinds(rect, collisionRects);
        const allowNoteBodyNudge: boolean = !stacked;
        return this.shouldTryFingeringHorizontalNudge(blockingKinds, allowNoteBodyNudge) ?
            this.getFingeringHorizontalNudgeOffsets(rect, stacked) : [0];
    }

    private getBestFingeringLabelCandidate(item: FingeringPlacementWorkItem, dy: number,
                                           collisionRects: FingeringCollisionRect[],
                                           stacked: boolean, measure: GraphicalMeasure,
                                           line: StaffLine): FingeringLabelCandidate {
        const shiftedRect: FingeringCollisionRect = this.shiftFingeringRect(item.initialRect, 0, dy);
        const horizontalOffsets: number[] =
            this.getFingeringCandidateHorizontalOffsets(shiftedRect, collisionRects, stacked);
        let bestCandidate: FingeringLabelCandidate;
        for (const dx of horizontalOffsets) {
            const snappedX: number = this.snapFingeringXToMeasure(item.label, item.ownerCenterX + dx, measure, line);
            const effectiveDx: number = snappedX - item.ownerCenterX;
            const rect: FingeringCollisionRect = this.shiftFingeringRect(item.initialRect, effectiveDx, dy);
            const cost: number = this.getFingeringCollisionCost(rect, collisionRects, stacked) +
                this.getFingeringHorizontalMovementCost(effectiveDx, stacked);
            if (!bestCandidate || cost < bestCandidate.cost - 0.0001 ||
                (Math.abs(cost - bestCandidate.cost) <= 0.0001 && Math.abs(effectiveDx) < Math.abs(bestCandidate.dx))) {
                    bestCandidate = { dx: effectiveDx, dy, rect, cost };
            }
        }
        return bestCandidate;
    }

    private getFingeringStackOverlapCost(candidates: FingeringLabelCandidate[]): number {
        let cost: number = 0;
        for (let i: number = 0; i < candidates.length; i++) {
            for (let j: number = i + 1; j < candidates.length; j++) {
                const actualOverlap: number = this.getRectOverlapArea(candidates[i].rect, candidates[j].rect);
                const paddedOverlap: number = this.getRectOverlapArea(
                    this.expandFingeringRect(candidates[i].rect, 0.025),
                    this.expandFingeringRect(candidates[j].rect, 0.025)
                );
                cost += actualOverlap * 10000 + paddedOverlap * 800;
            }
        }
        return cost;
    }

    private getFingeringStackAlignmentCost(candidates: FingeringLabelCandidate[]): number {
        if (candidates.length < 2) {
            return 0;
        }
        const meanDx: number = candidates.reduce(
            (sum: number, candidate: FingeringLabelCandidate) => sum + candidate.dx, 0) / candidates.length;
        return candidates.reduce((cost: number, candidate: FingeringLabelCandidate) =>
            cost + Math.abs(candidate.dx) * 26 + Math.abs(candidate.dx - meanDx) * 32, 0);
    }

    private getFingeringStackHardCollisionCost(candidates: FingeringLabelCandidate[],
                                               collisionRects: FingeringCollisionRect[]): number {
        if (candidates.length < 2) {
            return 0;
        }
        const envelopeRect: FingeringCollisionRect = this.getFingeringCandidateEnvelopeRect(candidates);
        let cost: number = 0;
        for (const collisionRect of collisionRects) {
            const kind: string = collisionRect.kind ?? "unknown";
            if (!this.isFingeringHardCollisionKind(kind)) {
                continue;
            }
            const padding: number = this.getFingeringCollisionClearancePadding(kind);
            const envelopeOverlap: number = this.getRectOverlapArea(
                this.expandFingeringRect(envelopeRect, padding),
                this.expandFingeringRect(collisionRect, padding)
            );
            if (envelopeOverlap <= 0) {
                continue;
            }
            const actualOverlap: number = this.getRectOverlapArea(envelopeRect, collisionRect);
            const kindWeight: number = this.getFingeringCollisionKindWeight(kind, true);
            const rectPenalty: number = collisionRect.penalty ?? 1;
            const softFactor: number = collisionRect.soft ? 0.45 : 1;
            cost += (actualOverlap * 90 + envelopeOverlap * this.getFingeringCollisionClearanceFactor(kind)) *
                kindWeight * rectPenalty * softFactor;
        }
        return cost;
    }

    private hasFingeringHardCollision(rect: FingeringCollisionRect,
                                      collisionRects: FingeringCollisionRect[]): boolean {
        for (const collisionRect of collisionRects) {
            if (!this.isFingeringHardCollisionKind(collisionRect.kind ?? "unknown")) {
                continue;
            }
            if (this.getRectOverlapArea(rect, collisionRect) > 0) {
                return true;
            }
        }
        return false;
    }

    private getFingeringResidualVerticalOffsets(items: FingeringPlacementWorkItem[],
                                                hardCollisionRects: FingeringCollisionRect[]): number[] {
        const placement: PlacementEnum = items[0].placement;
        const offsets: number[] = [0];
        for (const item of items) {
            if (!item.placedRect) {
                continue;
            }
            this.addFingeringTargetedVerticalOffsets(offsets, item.placedRect, placement, hardCollisionRects);
        }
        if (items.length > 1) {
            const envelopeRect: FingeringCollisionRect = this.getFingeringPlacedGroupEnvelopeRect(items);
            if (envelopeRect) {
                this.addFingeringTargetedVerticalOffsets(offsets, envelopeRect, placement, hardCollisionRects);
            }
        }
        for (let shift: number = -6.2; shift <= 6.2001; shift += 0.08) {
            offsets.push(shift);
        }

        const uniqueOffsets: number[] = [];
        const seen: Set<string> = new Set<string>();
        for (const offset of offsets) {
            const key: string = offset.toFixed(3);
            if (seen.has(key)) {
                continue;
            }
            seen.add(key);
            uniqueOffsets.push(offset);
        }
        return uniqueOffsets.sort((left: number, right: number) =>
            Math.abs(left) - Math.abs(right) || left - right);
    }

    private getFingeringResidualGroupCost(items: FingeringPlacementWorkItem[], dy: number,
                                          collisionRects: FingeringCollisionRect[]): {
        candidates: FingeringLabelCandidate[];
        cost: number;
    } {
        const stacked: boolean = items.length > 1;
        const candidates: FingeringLabelCandidate[] = items.map((item: FingeringPlacementWorkItem) => {
            const rect: FingeringCollisionRect = this.shiftFingeringRect(item.placedRect, 0, dy);
            return {
                dx: 0,
                dy,
                rect,
                cost: this.getFingeringCollisionCost(rect, collisionRects, stacked)
            };
        });
        const collisionCost: number = candidates.reduce(
            (sum: number, candidate: FingeringLabelCandidate) => sum + candidate.cost, 0);
        const movementCost: number = Math.abs(dy) * 22 + dy * dy * 10;
        return {
            candidates,
            cost: collisionCost + movementCost + this.getFingeringStackHardCollisionCost(candidates, collisionRects)
        };
    }

    private getFingeringHardCollisionOverlapArea(candidates: FingeringLabelCandidate[],
                                                 hardCollisionRects: FingeringCollisionRect[]): number {
        let overlapArea: number = 0;
        for (const candidate of candidates) {
            for (const collisionRect of hardCollisionRects) {
                overlapArea += this.getRectOverlapArea(candidate.rect, collisionRect);
            }
        }
        return overlapArea;
    }

    private resolveResidualFingeringHardCollisions(items: FingeringPlacementWorkItem[],
                                                   collisionRects: FingeringCollisionRect[]): void {
        const hardCollisionRects: FingeringCollisionRect[] = collisionRects.filter((rect: FingeringCollisionRect) =>
            this.isFingeringHardCollisionKind(rect.kind ?? "unknown"));
        if (hardCollisionRects.length === 0 ||
            !items.some((item: FingeringPlacementWorkItem) =>
                item.placedRect && this.hasFingeringHardCollision(item.placedRect, hardCollisionRects))) {
                    return;
        }

        let bestDy: number = 0;
        let bestCost: number = Number.POSITIVE_INFINITY;
        let bestHardOverlapArea: number = Number.POSITIVE_INFINITY;
        for (const dy of this.getFingeringResidualVerticalOffsets(items, hardCollisionRects)) {
            const resolution: { candidates: FingeringLabelCandidate[], cost: number } =
                this.getFingeringResidualGroupCost(items, dy, collisionRects);
            const hardOverlapArea: number =
                this.getFingeringHardCollisionOverlapArea(resolution.candidates, hardCollisionRects);
            if (hardOverlapArea < bestHardOverlapArea - 0.0001 ||
                (Math.abs(hardOverlapArea - bestHardOverlapArea) <= 0.0001 &&
                    (resolution.cost < bestCost - 0.0001 ||
                        (Math.abs(resolution.cost - bestCost) <= 0.0001 && Math.abs(dy) < Math.abs(bestDy))))) {
                    bestDy = dy;
                    bestCost = resolution.cost;
                    bestHardOverlapArea = hardOverlapArea;
            }
        }

        if (Math.abs(bestDy) <= 0.0001) {
            return;
        }
        for (const item of items) {
            const currentPosition: PointF2D = item.label.PositionAndShape.RelativePosition;
            this.setFingeringLabelPosition(
                item.label,
                currentPosition.x,
                currentPosition.y + bestDy,
                item.placement
            );
            item.placedRect = this.getLabelRect(item.label);
        }
    }

    private getFingeringGroupResolution(items: FingeringPlacementWorkItem[],
                                        collisionRects: FingeringCollisionRect[],
                                        measure: GraphicalMeasure, line: StaffLine): FingeringGroupResolution {
        const placement: PlacementEnum = items[0].placement;
        const stacked: boolean = items.length > 1;
        const horizontalSearchRadius: number = Math.max(...items.map((item: FingeringPlacementWorkItem) =>
            Math.max(...this.getFingeringHorizontalNudgeOffsets(item.initialRect, stacked)
                .map((offset: number) => Math.abs(offset)))));
        const relevantCollisionRects: FingeringCollisionRect[] =
            this.getRelevantFingeringGroupCollisionRects(items, collisionRects, horizontalSearchRadius);
        let bestResolution: FingeringGroupResolution;
        for (const dy of this.getFingeringVerticalSearchOffsetsForCollisions(items, placement, relevantCollisionRects)) {
            const candidates: FingeringLabelCandidate[] = items.map((item: FingeringPlacementWorkItem) =>
                this.getBestFingeringLabelCandidate(item, dy, relevantCollisionRects, stacked, measure, line));
            const collisionCost: number = candidates.reduce(
                (sum: number, candidate: FingeringLabelCandidate) => sum + candidate.cost, 0);
            const cost: number = collisionCost +
                this.getFingeringVerticalMovementCost(dy, items.length) +
                this.getFingeringStackOverlapCost(candidates) +
                this.getFingeringStackAlignmentCost(candidates) +
                this.getFingeringStackHardCollisionCost(candidates, relevantCollisionRects);
            if (!bestResolution || cost < bestResolution.cost - 0.0001 ||
                (Math.abs(cost - bestResolution.cost) <= 0.0001 && Math.abs(dy) < Math.abs(bestResolution.dy))) {
                    bestResolution = { dy, candidates, cost };
            }
        }
        return bestResolution;
    }

    private getFingeringDebugScope(): any {
        if (typeof globalThis === "undefined") {
            return undefined;
        }
        const scope: any = globalThis as any;
        return scope?.document ? scope : undefined;
    }

    private resetFingeringDebugRecords(): void {
        this.fingeringInternalDebugBoxesByKey.clear();
        const scope: any = this.getFingeringDebugScope();
        if (!scope) {
            return;
        }
        const generatedAt: string = new Date().toISOString();
        scope.__osmdFingeringDebug = {
            generatedAt,
            unitInPixels: 10,
            placements: []
        };
        scope.__osmdInternalCollisionDebug = {
            generatedAt,
            unitInPixels: 10,
            source: "fingering-calculator",
            boxes: []
        };
    }

    private getFingeringDebugSvgRect(rect: FingeringCollisionRect, line: StaffLine): FingeringDebugSvgRect {
        const unitInPixels: number = 10;
        const lineAbsolutePosition: PointF2D = line.PositionAndShape.AbsolutePosition ?? new PointF2D(0, 0);
        return {
            x: (lineAbsolutePosition.x + rect.left) * unitInPixels,
            y: (lineAbsolutePosition.y + rect.top) * unitInPixels,
            width: (rect.right - rect.left) * unitInPixels,
            height: (rect.bottom - rect.top) * unitInPixels
        };
    }

    private toFingeringDebugRect(rect: FingeringCollisionRect, line: StaffLine, index?: number,
                                 initialLabelRect?: FingeringCollisionRect,
                                 placedLabelRect?: FingeringCollisionRect): FingeringDebugRect {
        return {
            ...rect,
            index,
            kind: rect.kind ?? "unknown",
            svgRect: this.getFingeringDebugSvgRect(rect, line),
            overlapsInitialLabel: initialLabelRect ? this.getRectOverlapArea(rect, initialLabelRect) > 0 : undefined,
            overlapsPlacedLabel: placedLabelRect ? this.getRectOverlapArea(rect, placedLabelRect) > 0 : undefined
        };
    }

    private getInternalCollisionDebugKind(kind: string): InternalCollisionDebugKind {
        switch (kind) {
            case "beam":
            case "stem":
            case "note":
            case "notehead":
            case "existing-fingering":
            case "tie":
            case "ornament":
            case "articulation":
            case "ornament-reserve":
            case "rest":
                return kind as InternalCollisionDebugKind;
            default:
                return undefined;
        }
    }

    private getRoundedInternalCollisionCoordinate(value: number): string {
        return Number.isFinite(value) ? value.toFixed(3) : "";
    }

    private getInternalCollisionDebugKey(kind: InternalCollisionDebugKind, measure: number, staffId: number,
                                         rect: FingeringCollisionRect, debugKey?: string): string {
        return [
            measure,
            staffId,
            kind,
            debugKey ?? "",
            this.getRoundedInternalCollisionCoordinate(rect.left),
            this.getRoundedInternalCollisionCoordinate(rect.right),
            this.getRoundedInternalCollisionCoordinate(rect.top),
            this.getRoundedInternalCollisionCoordinate(rect.bottom)
        ].join("|");
    }

    private getInternalCollisionDebugReason(overlapsInitialLabel: boolean, overlapsPlacedLabel: boolean): string {
        const reasons: string[] = [];
        if (overlapsInitialLabel) {
            reasons.push("overlaps-initial-label");
        }
        if (overlapsPlacedLabel) {
            reasons.push("overlaps-placed-label");
        }
        return reasons.join(",");
    }

    private getFingeringPlacementDebugName(placement: PlacementEnum): string {
        return PlacementEnum[placement] ?? `${placement}`;
    }

    private addInternalCollisionDebugBox(debugState: any, box: InternalCollisionDebugBox): void {
        const existing: InternalCollisionDebugBox = this.fingeringInternalDebugBoxesByKey.get(box.key);
        if (existing) {
            existing.collision = existing.collision || box.collision;
            if (box.reason) {
                const reasons: Set<string> = new Set<string>(
                    `${existing.reason ?? ""},${box.reason}`.split(",").filter((reason: string) => reason.length > 0)
                );
                existing.reason = Array.from(reasons).join(",");
            }
            return;
        }

        this.fingeringInternalDebugBoxesByKey.set(box.key, box);
        debugState.boxes.push(box);
    }

    private toInternalCollisionDebugBox(rect: FingeringCollisionRect, kind: InternalCollisionDebugKind,
                                        line: StaffLine, fallbackMeasure: GraphicalMeasure,
                                        fingering: TechnicalInstruction, placement: PlacementEnum,
                                        stackIndex: number, initialLabelRect: FingeringCollisionRect,
                                        placedLabelRect: FingeringCollisionRect): InternalCollisionDebugBox {
        const measureNumber: number = rect.debugMeasure ?? fallbackMeasure.MeasureNumber;
        const staffId: number = rect.debugStaffId ?? line.ParentStaff?.Id ?? 0;
        const measureX: number = rect.debugMeasureX ?? fallbackMeasure.PositionAndShape.RelativePosition.x;
        const overlapsInitialLabel: boolean = this.getRectOverlapArea(rect, initialLabelRect) > 0;
        const overlapsPlacedLabel: boolean = this.getRectOverlapArea(rect, placedLabelRect) > 0;
        const reason: string = this.getInternalCollisionDebugReason(overlapsInitialLabel, overlapsPlacedLabel);
        return {
            key: this.getInternalCollisionDebugKey(kind, measureNumber, staffId, rect, rect.debugKey),
            kind,
            measure: measureNumber,
            staffId,
            measureX,
            left: rect.left,
            right: rect.right,
            top: rect.top,
            bottom: rect.bottom,
            source: "fingering-calculator",
            collision: overlapsInitialLabel || overlapsPlacedLabel || undefined,
            reason: reason || undefined,
            fingering: fingering.value,
            placement: this.getFingeringPlacementDebugName(placement),
            stackIndex,
            svgRect: this.getFingeringDebugSvgRect(rect, line)
        };
    }

    private addInternalLabelDebugBox(debugState: any, kind: "label-initial" | "label-placed",
                                     rect: FingeringCollisionRect, line: StaffLine, measure: GraphicalMeasure,
                                     fingering: TechnicalInstruction, placement: PlacementEnum, stackIndex: number,
                                     collision: boolean): void {
        const staffId: number = line.ParentStaff?.Id ?? 0;
        const key: string = this.getInternalCollisionDebugKey(
            kind,
            measure.MeasureNumber,
            staffId,
            rect,
            `fingering-label:${fingering.value}:${stackIndex}`
        );
        this.addInternalCollisionDebugBox(debugState, {
            key,
            kind,
            measure: measure.MeasureNumber,
            staffId,
            measureX: measure.PositionAndShape.RelativePosition.x,
            left: rect.left,
            right: rect.right,
            top: rect.top,
            bottom: rect.bottom,
            source: "fingering-calculator",
            collision: collision || undefined,
            reason: collision ? "label-overlap" : undefined,
            fingering: fingering.value,
            placement: this.getFingeringPlacementDebugName(placement),
            stackIndex,
            svgRect: this.getFingeringDebugSvgRect(rect, line)
        });
    }

    private recordInternalCollisionDebugBoxes(fingering: TechnicalInstruction, measure: GraphicalMeasure,
                                              line: StaffLine, placement: PlacementEnum, stackIndex: number,
                                              collisionRects: FingeringCollisionRect[],
                                              initialLabelRect: FingeringCollisionRect,
                                              placedLabelRect: FingeringCollisionRect): void {
        const scope: any = this.getFingeringDebugScope();
        const debugState: any = scope?.__osmdInternalCollisionDebug;
        if (!debugState?.boxes) {
            return;
        }

        let initialLabelCollides: boolean = false;
        let placedLabelCollides: boolean = false;
        for (const rect of collisionRects) {
            const kind: InternalCollisionDebugKind = this.getInternalCollisionDebugKind(rect.kind);
            if (!kind) {
                continue;
            }
            if (this.getRectOverlapArea(rect, initialLabelRect) > 0) {
                initialLabelCollides = true;
            }
            if (this.getRectOverlapArea(rect, placedLabelRect) > 0) {
                placedLabelCollides = true;
            }
            this.addInternalCollisionDebugBox(
                debugState,
                this.toInternalCollisionDebugBox(
                    rect,
                    kind,
                    line,
                    measure,
                    fingering,
                    placement,
                    stackIndex,
                    initialLabelRect,
                    placedLabelRect
                )
            );
        }

        if (initialLabelCollides) {
            this.addInternalLabelDebugBox(
                debugState,
                "label-initial",
                initialLabelRect,
                line,
                measure,
                fingering,
                placement,
                stackIndex,
                true
            );
        }
        this.addInternalLabelDebugBox(
            debugState,
            "label-placed",
            placedLabelRect,
            line,
            measure,
            fingering,
            placement,
            stackIndex,
            placedLabelCollides
        );
    }

    private recordFingeringDebugPlacement(label: GraphicalLabel, fingering: TechnicalInstruction, measure: GraphicalMeasure,
                                          line: StaffLine, system: MusicSystem, placement: PlacementEnum,
                                          stackIndex: number, collisionRects: FingeringCollisionRect[],
                                          initialLabelRect: FingeringCollisionRect,
                                          placedLabelRect: FingeringCollisionRect): void {
        const scope: any = this.getFingeringDebugScope();
        const debugState: any = scope?.__osmdFingeringDebug;
        if (!debugState?.placements) {
            return;
        }
        const measureNumber: number = measure.MeasureNumber;
        const record: any = {
            measure: measureNumber,
            measureX: measure.PositionAndShape.RelativePosition.x,
            staffId: line.ParentStaff?.Id,
            systemIndex: this.musicSystems.indexOf(system),
            staffLineIndex: system.StaffLines.indexOf(line),
            fingering: fingering.value,
            placement: this.getFingeringPlacementDebugName(placement),
            stackIndex,
            labelInitial: this.toFingeringDebugRect(initialLabelRect, line),
            labelPlaced: this.toFingeringDebugRect(placedLabelRect, line),
            collisionRects: collisionRects.map((rect: FingeringCollisionRect, index: number) =>
                this.toFingeringDebugRect(rect, line, index, initialLabelRect, placedLabelRect))
        };
        debugState.placements.push(record);
        this.recordInternalCollisionDebugBoxes(
            fingering,
            measure,
            line,
            placement,
            stackIndex,
            collisionRects,
            initialLabelRect,
            placedLabelRect
        );
    }

    private createFingeringPlacementWorkItem(label: GraphicalLabel, fingering: TechnicalInstruction,
                                             gse: GraphicalStaffEntry, measure: GraphicalMeasure,
                                             line: StaffLine, stackIndex: number): FingeringPlacementWorkItem {
        const placement: PlacementEnum = this.getFingeringPlacement(measure, fingering);
        const owningNote: GraphicalNote = this.getFingeringOwnerNote(gse, fingering);
        const anchorY: number = this.getFingeringAnchorY(placement);
        const staffEntryPositionX: number =
            gse.PositionAndShape.RelativePosition.x + measure.PositionAndShape.RelativePosition.x;
        const labelHeight: number = this.getFingeringLabelHeight(label);
        const ownerRect: FingeringCollisionRect = owningNote ?
            this.getBoundingBoxRectInStaffLine(owningNote.PositionAndShape, line, false) : undefined;
        const ownerCenterX: number = this.snapFingeringXToMeasure(
            label,
            ownerRect ? (ownerRect.left + ownerRect.right) / 2 : staffEntryPositionX,
            measure,
            line
        );
        const usableAnchorY: number = Number.isFinite(anchorY) ? anchorY :
            (placement === PlacementEnum.Above ? 0 : this.rules.StaffHeight);
        const distance: number = 0.18 + stackIndex * this.getFingeringStackSpacing(labelHeight);
        const y: number = placement === PlacementEnum.Above ? usableAnchorY - distance : usableAnchorY + distance;
        this.setFingeringLabelPosition(label, ownerCenterX, y, placement);
        return {
            fingering,
            label,
            placement,
            stackIndex,
            ownerCenterX,
            initialY: y,
            labelHeight,
            initialRect: this.getLabelRect(label)
        };
    }

    private getFingeringGroupDebugCollisionRects(items: FingeringPlacementWorkItem[],
                                                 currentItem: FingeringPlacementWorkItem,
                                                 collisionRects: FingeringCollisionRect[],
                                                 measure: GraphicalMeasure,
                                                 line: StaffLine): FingeringCollisionRect[] {
        const siblingRects: FingeringCollisionRect[] = [];
        for (let itemIndex: number = 0; itemIndex < items.length; itemIndex++) {
            const item: FingeringPlacementWorkItem = items[itemIndex];
            if (item === currentItem || !item.placedRect) {
                continue;
            }
            this.addRectIfValid(siblingRects, this.withFingeringDebugSource({
                ...item.placedRect,
                isFingering: true,
                kind: "existing-fingering"
            }, measure, line, `stack-fingering:${itemIndex}`));
        }
        return collisionRects.concat(siblingRects);
    }

    private shiftFingeringPlacementWorkItem(item: FingeringPlacementWorkItem, dy: number): void {
        item.initialY += dy;
        this.setFingeringLabelPosition(item.label, item.ownerCenterX, item.initialY, item.placement);
        item.initialRect = this.getLabelRect(item.label);
    }

    private getFingeringHorizontalOverlap(a: FingeringCollisionRect, b: FingeringCollisionRect): number {
        return Math.min(a.right, b.right) - Math.max(a.left, b.left);
    }

    private separateOverlappingFingeringStack(items: FingeringPlacementWorkItem[]): void {
        if (items.length < 2) {
            return;
        }
        const placement: PlacementEnum = items[0].placement;
        const outwardSign: number = placement === PlacementEnum.Above ? -1 : 1;
        const stackGap: number = 0.04;
        const sortedItems: FingeringPlacementWorkItem[] = items.slice().sort(
            (a: FingeringPlacementWorkItem, b: FingeringPlacementWorkItem) => a.stackIndex - b.stackIndex);
        for (let itemIndex: number = 1; itemIndex < sortedItems.length; itemIndex++) {
            const item: FingeringPlacementWorkItem = sortedItems[itemIndex];
            let requiredShift: number = 0;
            for (let previousIndex: number = 0; previousIndex < itemIndex; previousIndex++) {
                const previousItem: FingeringPlacementWorkItem = sortedItems[previousIndex];
                if (this.getFingeringHorizontalOverlap(item.initialRect, previousItem.initialRect) <= 0) {
                    continue;
                }
                const overlap: number = placement === PlacementEnum.Above ?
                    item.initialRect.bottom - previousItem.initialRect.top + stackGap :
                    previousItem.initialRect.bottom - item.initialRect.top + stackGap;
                requiredShift = Math.max(requiredShift, overlap);
            }
            if (requiredShift <= 0) {
                continue;
            }
            for (let shiftedIndex: number = itemIndex; shiftedIndex < sortedItems.length; shiftedIndex++) {
                this.shiftFingeringPlacementWorkItem(sortedItems[shiftedIndex], requiredShift * outwardSign);
            }
        }
    }

    private placeFingeringLabelGroup(items: FingeringPlacementWorkItem[], gse: GraphicalStaffEntry,
                                     measure: GraphicalMeasure, line: StaffLine, system: MusicSystem): void {
        if (items.length === 0) {
            return;
        }
        this.separateOverlappingFingeringStack(items);
        const placement: PlacementEnum = items[0].placement;
        const collisionRects: FingeringCollisionRect[] = this.getFingeringCollisionRects(system, line, measure, placement);
        const resolution: FingeringGroupResolution = this.getFingeringGroupResolution(items, collisionRects, measure, line);
        for (let itemIndex: number = 0; itemIndex < items.length; itemIndex++) {
            const item: FingeringPlacementWorkItem = items[itemIndex];
            const candidate: FingeringLabelCandidate = resolution.candidates[itemIndex];
            const resolvedX: number = this.snapFingeringXToMeasure(item.label, item.ownerCenterX + candidate.dx, measure, line);
            const resolvedY: number = item.initialY + candidate.dy;
            this.setFingeringLabelPosition(item.label, resolvedX, resolvedY, item.placement);
            item.placedRect = this.getLabelRect(item.label);
        }
        this.resolveResidualFingeringHardCollisions(items, collisionRects);

        for (const item of items) {
            const debugCollisionRects: FingeringCollisionRect[] =
                this.getFingeringGroupDebugCollisionRects(items, item, collisionRects, measure, line);
            this.recordFingeringDebugPlacement(
                item.label,
                item.fingering,
                measure,
                line,
                system,
                item.placement,
                item.stackIndex,
                debugCollisionRects,
                item.initialRect,
                item.placedRect
            );
        }
    }

    private updateFingeringSkyBottomLine(label: GraphicalLabel, placement: PlacementEnum, line: StaffLine): void {
        const labelRect: FingeringCollisionRect = placement === PlacementEnum.Below ?
            this.getCompactFingeringStackRect(this.getLabelRect(label)) :
            this.getLabelRect(label);
        if (placement === PlacementEnum.Above) {
            line.SkyBottomLineCalculator.updateSkyLineInRange(labelRect.left, labelRect.right, labelRect.top);
        } else if (placement === PlacementEnum.Below) {
            line.SkyBottomLineCalculator.updateBottomLineInRange(labelRect.left, labelRect.right, labelRect.bottom);
        }
    }

    public calculateFingerings(): void {
        if (this.rules.FingeringPosition === PlacementEnum.Left ||
            this.rules.FingeringPosition === PlacementEnum.Right) {
                return;
        }
        this.resetFingeringDebugRecords();
        for (const system of this.musicSystems) {
            for (const line of system.StaffLines) {
                for (const measure of line.Measures) {
                    if (measure.isTabMeasure && !this.rules.TabFingeringsRendered) {
                        continue; // don't duplicate fingerings into tab measures. tab notes are already
                    }
                    const defaultPlacement: PlacementEnum = this.getFingeringPlacement(measure);
                    for (const gse of measure.staffEntries) {
                        gse.FingeringEntries = [];
                        const fingerings: TechnicalInstruction[] = [];
                        for (const voiceEntry of gse.graphicalVoiceEntries) {
                            if (voiceEntry.parentVoiceEntry.IsGrace) {
                                continue;
                            }
                            // Sibelius: can have multiple fingerings per note, so we need to check voice entry instructions, not note.Fingering
                            for (const instruction of voiceEntry.parentVoiceEntry.TechnicalInstructions) {
                                if (instruction.type === TechnicalInstructionType.Fingering) {
                                    fingerings.push(instruction);
                                }
                            }
                            // for (const note of voiceEntry.notes) {
                            //     const sourceNote: Note = note.sourceNote;
                            //     if (sourceNote.Fingering && !sourceNote.IsGraceNote) {
                            //         fingerings.push(sourceNote.Fingering);
                            //     }
                            // }
                        }
                        const fingeringOwners: GraphicalNote[] = fingerings.map(
                            (fingering: TechnicalInstruction) => this.getFingeringOwnerNote(gse, fingering));
                        const stackSlots: number[] = this.getFingeringStackSlots(fingerings, fingeringOwners, gse, defaultPlacement);
                        const placementItems: { fingering: TechnicalInstruction, slot: number, sourceIndex: number }[] =
                            fingerings.map((fingering: TechnicalInstruction, index: number) => ({
                                fingering,
                                slot: stackSlots[index],
                                sourceIndex: index
                            })).sort((a, b) => a.slot - b.slot || a.sourceIndex - b.sourceIndex);

                        const placementGroups: { placement: PlacementEnum, items: FingeringPlacementWorkItem[] }[] = [];
                        for (const placementItem of placementItems) {
                            const fingering: TechnicalInstruction = placementItem.fingering;
                            const placement: PlacementEnum = this.getFingeringPlacement(measure, fingering);
                            const alignment: TextAlignmentEnum =
                                placement === PlacementEnum.Above ? TextAlignmentEnum.CenterBottom : TextAlignmentEnum.CenterTop;
                            const label: Label = new Label(fingering.value, alignment);
                            label.fontFamily = fingering.fontFamily || this.rules.FingeringFontFamily;
                            label.fontStyle = FontStyles.Regular;
                            const gLabel: GraphicalLabel = new GraphicalLabel(
                                label, this.rules.FingeringTextSize, label.textAlignment, this.rules, line.PositionAndShape);
                            const workItem: FingeringPlacementWorkItem =
                                this.createFingeringPlacementWorkItem(gLabel, fingering, gse, measure, line, placementItem.slot);
                            let placementGroup: { placement: PlacementEnum, items: FingeringPlacementWorkItem[] } =
                                placementGroups.find((group: { placement: PlacementEnum, items: FingeringPlacementWorkItem[] }) =>
                                    group.placement === workItem.placement);
                            if (!placementGroup) {
                                placementGroup = { placement: workItem.placement, items: [] };
                                placementGroups.push(placementGroup);
                            }
                            placementGroup.items.push(workItem);
                        }

                        for (const placementGroup of placementGroups) {
                            this.placeFingeringLabelGroup(placementGroup.items, gse, measure, line, system);
                            for (const item of placementGroup.items) {
                                gse.FingeringEntries.push(item.label);
                                this.updateFingeringSkyBottomLine(item.label, item.placement, line);
                            }
                        }
                    }
                }
            }
        }
    }

    private optimizeRestPlacement(): void {
        for (let idx2: number = 0, len2: number = this.musicSystems.length; idx2 < len2; ++idx2) {
            const system: MusicSystem = this.musicSystems[idx2];
            for (let idx3: number = 0, len3: number = system.StaffLines.length; idx3 < len3; ++idx3) {
                const line: StaffLine = system.StaffLines[idx3];
                for (let idx4: number = 0, len4: number = line.Measures.length; idx4 < len4; ++idx4) {
                    const measure: GraphicalMeasure = line.Measures[idx4];
                    for (let idx5: number = 0, len5: number = measure.staffEntries.length; idx5 < len5; ++idx5) {
                        const graphicalStaffEntry: GraphicalStaffEntry = measure.staffEntries[idx5];
                        this.optimizeRestNotePlacement(graphicalStaffEntry, measure);
                    }
                }
            }
        }
    }

    private collectRestPlacementNotes(graphicalStaffEntry: GraphicalStaffEntry,
                                      restNotes: GraphicalNote[],
                                      pitchedNotes: GraphicalNote[]): void {
        for (const voiceEntry of graphicalStaffEntry.graphicalVoiceEntries) {
            for (const graphicalNote of voiceEntry.notes) {
                if (graphicalNote.sourceNote?.isRest()) {
                    restNotes.push(graphicalNote);
                } else if (graphicalNote.sourceNote?.PrintObject !== false) {
                    pitchedNotes.push(graphicalNote);
                }
            }
        }
    }

    private collectRestPlacementMeasurePitchedNotes(measure: GraphicalMeasure,
                                                    seedNotes: GraphicalNote[]): GraphicalNote[] {
        const notes: GraphicalNote[] = seedNotes.slice();
        for (const staffEntry of measure.staffEntries ?? []) {
            for (const voiceEntry of staffEntry.graphicalVoiceEntries ?? []) {
                for (const graphicalNote of voiceEntry.notes ?? []) {
                    if (graphicalNote.sourceNote?.isRest() || graphicalNote.sourceNote?.PrintObject === false ||
                        notes.indexOf(graphicalNote) >= 0) {
                            continue;
                    }
                    notes.push(graphicalNote);
                }
            }
        }
        return notes;
    }

    private getRestVexFlowNote(restNote: GraphicalNote): any {
        return (restNote as any).vfnote?.[0];
    }

    private getRestVexFlowNoteIndex(restNote: GraphicalNote): number {
        const noteheadIndex: number = Number((restNote as any).vfnote?.[1] ?? (restNote as any).vfnoteIndex ?? 0);
        return Number.isFinite(noteheadIndex) ? noteheadIndex : 0;
    }

    private getVexFlowElementBoundingRectPx(element: any, padding: number = 0): FingeringPxRect {
        let box: any;
        try {
            box = typeof element?.getBoundingBox === "function" ? element.getBoundingBox() : element?.boundingBox;
        } catch (e) {
            return undefined;
        }
        if (!box) {
            return undefined;
        }
        const x: number = Number(typeof box.getX === "function" ? box.getX() : box.x);
        const y: number = Number(typeof box.getY === "function" ? box.getY() : box.y);
        const width: number = Number(typeof box.getW === "function" ? box.getW() : box.w ?? box.width);
        const height: number = Number(typeof box.getH === "function" ? box.getH() : box.h ?? box.height);
        if (!Number.isFinite(x) || !Number.isFinite(y) || !Number.isFinite(width) || !Number.isFinite(height) ||
            Math.abs(width) <= 0.01 || Math.abs(height) <= 0.01) {
                return undefined;
        }
        return this.normalizeFingeringPxRect({ x, y, width, height }, padding);
    }

    private getRestGlyphFallbackRectPx(restNote: GraphicalNote, vfNote: any): FingeringPxRect {
        if (!vfNote) {
            return undefined;
        }
        const index: number = this.getRestVexFlowNoteIndex(restNote);
        const stave: any = typeof vfNote.getStave === "function" ? vfNote.getStave() : vfNote.stave;
        const glyphMetrics: { width: number, height: number } =
            this.getFingeringVexFlowGlyphMetrics(vfNote.glyph, vfNote);
        const glyphWidth: number = Number(typeof vfNote.getGlyphWidth === "function" ? vfNote.getGlyphWidth() : glyphMetrics.width);
        const staffSpace: number = Number(stave?.getSpacingBetweenLines?.() ?? 10);
        const width: number = Math.min(Math.max(
            Number.isFinite(glyphWidth) ? glyphWidth : 0,
            glyphMetrics.width,
            staffSpace * 0.85,
            10
        ), staffSpace * 1.6);
        const height: number = Math.min(Math.max(
            glyphMetrics.height,
            staffSpace * 1.6,
            14
        ), staffSpace * 2.3);

        const xs: number[] = typeof vfNote.getXs === "function" ? vfNote.getXs() : [];
        const absoluteX: number = Number(typeof vfNote.getAbsoluteX === "function" ? vfNote.getAbsoluteX() : vfNote.getX?.());
        const keyProps: any[] = typeof vfNote.getKeyProps === "function" ? vfNote.getKeyProps() : vfNote.keyProps;
        const keyLine: number = Number(keyProps?.[index]?.line ?? keyProps?.[0]?.line);
        const fallbackY: number = Number.isFinite(keyLine) && typeof stave?.getYForNote === "function" ?
            Number(stave.getYForNote(keyLine)) : undefined;

        const centerX: number = Number.isFinite(xs?.[index]) ? xs[index] :
            Number.isFinite(absoluteX) ? absoluteX + width / 2 : undefined;
        const centerY: number = fallbackY;
        if (!Number.isFinite(centerX) || !Number.isFinite(centerY)) {
            return undefined;
        }
        return this.normalizeFingeringPxRect({
            x: centerX - width / 2,
            y: centerY - height / 2,
            width,
            height
        }, 0.4);
    }

    private getRestGlyphCollisionRect(restNote: GraphicalNote, measure: GraphicalMeasure): FingeringCollisionRect {
        const vfNote: any = this.getRestVexFlowNote(restNote);
        if (!vfNote) {
            return undefined;
        }
        const measureX: number = measure?.PositionAndShape?.RelativePosition?.x ?? 0;
        const staveOrigin: { x: number, y: number } = this.getVexFlowStaveOriginPx(vfNote);
        const pxRect: FingeringPxRect = this.getRestGlyphFallbackRectPx(restNote, vfNote) ??
            this.getVexFlowElementBoundingRectPx(vfNote, 1.0);
        const rect: FingeringCollisionRect =
            this.fingeringPxRectToCollisionRect(pxRect, measureX, staveOrigin, "rest", 1);
        if (!rect || rect.bottom < -4 || rect.top > this.rules.StaffHeight + 4) {
            return undefined;
        }
        return rect;
    }

    private getRestCollisionRects(restNote: GraphicalNote, measure: GraphicalMeasure,
                                  line: StaffLine): FingeringCollisionRect[] {
        const rects: FingeringCollisionRect[] = [];
        this.addRectIfValid(rects, this.getRestGlyphCollisionRect(restNote, measure));
        if (rects.length === 0 && line) {
            const restRect: FingeringCollisionRect =
                this.getBoundingBoxRectInStaffLine(restNote.PositionAndShape, line, true);
            if (restRect) {
                restRect.kind = "rest";
            }
            this.addRectIfValid(rects, restRect);
        }
        return rects;
    }

    private getRestCollisionNoteRects(graphicalNotes: GraphicalNote[], measure: GraphicalMeasure,
                                      line: StaffLine): FingeringCollisionRect[] {
        const rects: FingeringCollisionRect[] = [];
        if (!line) {
            return rects;
        }
        for (const graphicalNote of graphicalNotes) {
            if (graphicalNote.sourceNote?.isRest()) {
                continue;
            }
            this.addFingeringNoteheadCollisionRect(rects, graphicalNote, measure, line);
            this.addFingeringStemCollisionRect(rects, graphicalNote, measure, line);
        }
        return rects;
    }

    private getRestCollisionCost(restRects: FingeringCollisionRect[], noteRects: FingeringCollisionRect[]): number {
        let cost: number = 0;
        for (const restRect of restRects) {
            for (const noteRect of noteRects) {
                const restPadding: number = noteRect.kind === "stem" ? 0.015 : noteRect.kind === "rest" ? 0.16 : 0.1;
                const obstaclePadding: number = noteRect.kind === "stem" ? 0.005 : noteRect.kind === "rest" ? 0.16 : 0.06;
                const expandedRestRect: FingeringCollisionRect = this.expandFingeringRect(restRect, restPadding);
                const expandedNoteRect: FingeringCollisionRect = this.expandFingeringRect(noteRect, obstaclePadding);
                const overlap: number = this.getRectOverlapArea(expandedRestRect, expandedNoteRect);
                if (overlap <= 0) {
                    continue;
                }
                const kindWeight: number =
                    noteRect.kind === "notehead" ? 8 :
                    noteRect.kind === "rest" ? 9 :
                    noteRect.kind === "stem" ? 0.02 :
                    noteRect.kind === "note" ? 0.6 : 0.3;
                cost += overlap * kindWeight;
            }
        }
        return cost;
    }

    private getRestHardCollisionCost(restRects: FingeringCollisionRect[], noteRects: FingeringCollisionRect[]): number {
        let cost: number = 0;
        for (const restRect of restRects) {
            for (const noteRect of noteRects) {
                if (noteRect.kind !== "notehead" && noteRect.kind !== "rest") {
                    continue;
                }
                const overlap: number = this.getRectOverlapArea(
                    this.expandFingeringRect(restRect, 0.035),
                    this.expandFingeringRect(noteRect, noteRect.kind === "rest" ? 0.08 : 0.04)
                );
                if (overlap <= 0) {
                    continue;
                }
                cost += overlap * (noteRect.kind === "rest" ? 2 : 1);
            }
        }
        return cost;
    }

    private restNoteCollidesWithGraphicalNotes(restNote: GraphicalNote, graphicalNotes: GraphicalNote[],
                                               measure: GraphicalMeasure, line: StaffLine): boolean {
        const noteRects: FingeringCollisionRect[] = this.getRestCollisionNoteRects(graphicalNotes, measure, line);
        const restRects: FingeringCollisionRect[] = this.getRestCollisionRects(restNote, measure, line);
        if (this.getRestCollisionCost(restRects, noteRects) > 0) {
            return true;
        }
        for (const graphicalNote of graphicalNotes) {
            if (restNote.PositionAndShape.marginCollisionDetection(graphicalNote.PositionAndShape)) {
                return true;
            }
        }
        return false;
    }

    private getRestVoiceStemDirection(voiceEntry: VoiceEntry): number {
        switch (voiceEntry?.StemDirection) {
            case StemDirectionType.Up:
                return this.getFingeringVexFlowStemUp();
            case StemDirectionType.Down:
                return this.getFingeringVexFlowStemDown();
            default:
                break;
        }
        switch (voiceEntry?.WantedStemDirection) {
            case StemDirectionType.Up:
                return this.getFingeringVexFlowStemUp();
            case StemDirectionType.Down:
                return this.getFingeringVexFlowStemDown();
            default:
                break;
        }
        return undefined;
    }

    private getRestPitchedNoteStemDirection(graphicalNote: GraphicalNote): number {
        if (!graphicalNote || graphicalNote.sourceNote?.isRest()) {
            return undefined;
        }
        const vfNote: any = (graphicalNote as any).vfnote?.[0];
        const hasStem: boolean = typeof vfNote?.hasStem === "function" ? vfNote.hasStem() : !!vfNote?.getStem?.();
        if (hasStem || typeof vfNote?.hasStem !== "function") {
            const actualDirection: number = Number(
                typeof vfNote?.getStemDirection === "function" ? vfNote.getStemDirection() : vfNote?.stem_direction
            );
            if (Number.isFinite(actualDirection) && actualDirection !== 0) {
                return actualDirection;
            }
        }
        return this.getRestVoiceStemDirection(
            graphicalNote.sourceNote?.ParentVoiceEntry ?? graphicalNote.parentVoiceEntry?.parentVoiceEntry
        );
    }

    private getRestTargetGraphicalNote(restRect: FingeringCollisionRect, graphicalNotes: GraphicalNote[],
                                       measure: GraphicalMeasure, line: StaffLine): GraphicalNote {
        if (graphicalNotes.length === 0) {
            return undefined;
        }
        if (!restRect || !line) {
            return graphicalNotes[0];
        }
        const restCenterX: number = (restRect.left + restRect.right) / 2;
        const restCenterY: number = (restRect.top + restRect.bottom) / 2;
        const candidates: { note: GraphicalNote, rect: FingeringCollisionRect }[] = [];
        for (const graphicalNote of graphicalNotes) {
            const noteheadRect: FingeringCollisionRect =
                this.getFingeringNoteheadCollisionRect(graphicalNote, measure, line);
            if (noteheadRect) {
                candidates.push({ note: graphicalNote, rect: noteheadRect });
            }
        }
        if (candidates.length === 0) {
            return graphicalNotes[0];
        }
        return candidates.sort((a: { note: GraphicalNote, rect: FingeringCollisionRect },
                                b: { note: GraphicalNote, rect: FingeringCollisionRect }) => {
            const overlapA: number = Math.max(0, Math.min(restRect.right, a.rect.right) - Math.max(restRect.left, a.rect.left));
            const overlapB: number = Math.max(0, Math.min(restRect.right, b.rect.right) - Math.max(restRect.left, b.rect.left));
            const distanceXA: number = Math.abs(restCenterX - (a.rect.left + a.rect.right) / 2);
            const distanceXB: number = Math.abs(restCenterX - (b.rect.left + b.rect.right) / 2);
            const distanceYA: number = Math.abs(restCenterY - (a.rect.top + a.rect.bottom) / 2);
            const distanceYB: number = Math.abs(restCenterY - (b.rect.top + b.rect.bottom) / 2);
            return overlapB - overlapA || distanceXA - distanceXB || distanceYA - distanceYB;
        })[0].note;
    }

    private shouldPlaceRestAboveCollidingNote(restNote: GraphicalNote, targetNote: GraphicalNote,
                                              graphicalNotes: GraphicalNote[]): boolean {
        let collidingStemDirection: number = this.getRestPitchedNoteStemDirection(targetNote);
        if (!Number.isFinite(collidingStemDirection)) {
            collidingStemDirection = graphicalNotes
                .map((note: GraphicalNote) => this.getRestPitchedNoteStemDirection(note))
                .find((direction: number) => Number.isFinite(direction));
        }
        if (collidingStemDirection > 0) {
            return false;
        }
        if (collidingStemDirection < 0) {
            return true;
        }

        const restStemDirection: number = this.getRestVoiceStemDirection(restNote.sourceNote?.ParentVoiceEntry);
        if (restStemDirection > 0) {
            return true;
        }
        if (restStemDirection < 0) {
            return false;
        }
        const voiceId: number = restNote.sourceNote?.ParentVoiceEntry?.ParentVoice?.VoiceId ?? 0;
        if (voiceId > 0) {
            return voiceId % 2 === 1;
        }
        return true;
    }

    private clampRestCollisionLineShift(lineShift: number): number {
        if (!Number.isFinite(lineShift)) {
            return 0;
        }
        return Math.max(-3.5, Math.min(3.5, lineShift));
    }

    private clampRestCollisionKeyLine(line: number): number {
        return Math.max(-1.5, Math.min(5.5, line));
    }

    private getRestNoteLineShift(restNote: GraphicalNote): number {
        return this.clampRestCollisionLineShift(Number(restNote.lineShift ?? 0));
    }

    private setRestNoteLineShift(restNote: GraphicalNote, lineShift: number): void {
        const nextLineShift: number = this.clampRestCollisionLineShift(lineShift);
        const previousLineShift: number = this.getRestNoteLineShift(restNote);
        const delta: number = nextLineShift - previousLineShift;
        restNote.lineShift = nextLineShift;
        if (Math.abs(delta) <= 0.0001) {
            return;
        }
        const vfNote: any = this.getRestVexFlowNote(restNote);
        const keyProps: any[] = typeof vfNote?.getKeyProps === "function" ? vfNote.getKeyProps() : vfNote?.keyProps;
        const noteheadIndex: number = this.getRestVexFlowNoteIndex(restNote);
        const keyProp: any = keyProps?.[noteheadIndex] ?? keyProps?.[0];
        if (!keyProp || !Number.isFinite(Number(keyProp.line))) {
            return;
        }
        const nextLine: number = this.clampRestCollisionKeyLine(Number(keyProp.line) + delta);
        if (typeof vfNote.setKeyLine === "function") {
            vfNote.setKeyLine(noteheadIndex, nextLine);
        } else {
            keyProp.line = nextLine;
        }
    }

    private setRestCollisionCandidateState(restNote: GraphicalNote, basePosition: PointF2D,
                                           baseLineShift: number, candidateLineShift: number): void {
        const boundedLineShift: number = this.clampRestCollisionLineShift(candidateLineShift);
        this.setRestNoteLineShift(restNote, boundedLineShift);
        const lineDelta: number = boundedLineShift - baseLineShift;
        restNote.PositionAndShape.RelativePosition = new PointF2D(basePosition.x, basePosition.y - lineDelta);
    }

    private getRestLineShiftCandidates(baseLineShift: number): number[] {
        const candidates: number[] = [baseLineShift];
        for (let lineShift: number = -3.5; lineShift <= 3.5001; lineShift += 0.5) {
            candidates.push(lineShift);
        }
        const uniqueCandidates: number[] = Array.from(new Set(candidates.map((candidate: number) =>
            this.clampRestCollisionLineShift(candidate))));
        return uniqueCandidates.sort((a: number, b: number) =>
            Math.abs(a - baseLineShift) - Math.abs(b - baseLineShift) || a - b);
    }

    private getAdjacentRestPlacementScore(restRects: FingeringCollisionRect[], noteRects: FingeringCollisionRect[],
                                          targetNoteheadRect: FingeringCollisionRect, placeAbove: boolean,
                                          candidateLineShift: number, baseLineShift: number): number {
        const restRect: FingeringCollisionRect = restRects[0];
        if (!restRect) {
            return Number.POSITIVE_INFINITY;
        }
        const collisionCost: number = this.getRestCollisionCost(restRects, noteRects);
        const hardCollisionCost: number = this.getRestHardCollisionCost(restRects, noteRects);
        const lineDelta: number = candidateLineShift - baseLineShift;
        const preferredDirection: number = placeAbove ? 1 : -1;
        const wrongDirectionCost: number = Math.abs(lineDelta) > 0.0001 && Math.sign(lineDelta) !== preferredDirection ? 140 : 0;
        const movementCost: number = Math.pow(Math.abs(lineDelta), 1.28) * 48 + wrongDirectionCost;
        let adjacencyCost: number = 0;
        if (targetNoteheadRect) {
            const clearance: number = 0.1;
            const desiredEdge: number = placeAbove ? targetNoteheadRect.top - clearance : targetNoteheadRect.bottom + clearance;
            const actualEdge: number = placeAbove ? restRect.bottom : restRect.top;
            const wrongSideDistance: number = placeAbove ?
                Math.max(0, actualEdge - desiredEdge) :
                Math.max(0, desiredEdge - actualEdge);
            const edgeDistance: number = Math.abs(actualEdge - desiredEdge);
            const restCenterX: number = (restRect.left + restRect.right) / 2;
            const noteCenterX: number = (targetNoteheadRect.left + targetNoteheadRect.right) / 2;
            const xDistance: number = Math.abs(restCenterX - noteCenterX);
            adjacencyCost =
                wrongSideDistance * 5000 +
                edgeDistance * 80 +
                xDistance * 0.3;
        }
        return hardCollisionCost * 1000000 +
            collisionCost * 16000 +
            adjacencyCost +
            movementCost;
    }

    private resolveRestLineShiftCollision(restNote: GraphicalNote, targetGraphicalNotes: GraphicalNote[],
                                          collisionGraphicalNotes: GraphicalNote[],
                                          restBlockerRects: FingeringCollisionRect[],
                                          graphicalStaffEntry: GraphicalStaffEntry,
                                          measure: GraphicalMeasure, line: StaffLine): void {
        const noteRects: FingeringCollisionRect[] =
            this.getRestCollisionNoteRects(collisionGraphicalNotes, measure, line).concat(restBlockerRects);
        if (noteRects.length === 0) {
            return;
        }
        const initialRestRects: FingeringCollisionRect[] = this.getRestCollisionRects(restNote, measure, line);
        const targetNote: GraphicalNote = this.getRestTargetGraphicalNote(initialRestRects[0], targetGraphicalNotes, measure, line) ??
            this.getRestTargetGraphicalNote(initialRestRects[0], collisionGraphicalNotes, measure, line);
        const targetNoteheadRect: FingeringCollisionRect = targetNote ?
            this.getFingeringNoteheadCollisionRect(targetNote, measure, line) : undefined;
        const placeAbove: boolean = this.shouldPlaceRestAboveCollidingNote(restNote, targetNote, collisionGraphicalNotes);
        const basePosition: PointF2D = restNote.PositionAndShape.RelativePosition;
        const baseLineShift: number = this.getRestNoteLineShift(restNote);
        let bestLineShift: number = baseLineShift;
        let bestScore: number = Number.POSITIVE_INFINITY;

        for (const candidateLineShift of this.getRestLineShiftCandidates(baseLineShift)) {
            this.setRestCollisionCandidateState(restNote, basePosition, baseLineShift, candidateLineShift);
            graphicalStaffEntry.PositionAndShape.calculateAbsolutePositionsRecursiveWithoutTopelement();
            const restRects: FingeringCollisionRect[] = this.getRestCollisionRects(restNote, measure, line);
            const score: number = this.getAdjacentRestPlacementScore(
                restRects, noteRects, targetNoteheadRect, placeAbove, candidateLineShift, baseLineShift);
            if (score < bestScore - 0.0001 ||
                (Math.abs(score - bestScore) <= 0.0001 &&
                    Math.abs(candidateLineShift - baseLineShift) < Math.abs(bestLineShift - baseLineShift))) {
                    bestScore = score;
                    bestLineShift = candidateLineShift;
            }
        }

        this.setRestCollisionCandidateState(restNote, basePosition, baseLineShift, bestLineShift);
        graphicalStaffEntry.PositionAndShape.calculateAbsolutePositionsRecursiveWithoutTopelement();
    }

    private calculateRestNotesPlacementWithCollisionDetection(graphicalStaffEntry: GraphicalStaffEntry,
                                                              measure: GraphicalMeasure,
                                                              restNotes: GraphicalNote[],
                                                              graphicalNotes: GraphicalNote[]): void {
        const line: StaffLine = measure.ParentStaffLine;
        graphicalStaffEntry.PositionAndShape.calculateAbsolutePositionsRecursiveWithoutTopelement();
        const collisionGraphicalNotes: GraphicalNote[] =
            this.collectRestPlacementMeasurePitchedNotes(measure, graphicalNotes);
        const placedRestRects: FingeringCollisionRect[] = [];
        for (const restNote of restNotes) {
            const restRects: FingeringCollisionRect[] = this.getRestCollisionRects(restNote, measure, line);
            const collidesWithNotes: boolean =
                this.restNoteCollidesWithGraphicalNotes(restNote, collisionGraphicalNotes, measure, line);
            const collidesWithPlacedRests: boolean = this.getRestCollisionCost(restRects, placedRestRects) > 0;
            if (graphicalNotes.length > 0 || collidesWithNotes || collidesWithPlacedRests) {
                this.resolveRestLineShiftCollision(
                    restNote,
                    graphicalNotes,
                    collisionGraphicalNotes,
                    placedRestRects,
                    graphicalStaffEntry,
                    measure,
                    line
                );
            }
            const finalRestRects: FingeringCollisionRect[] = this.getRestCollisionRects(restNote, measure, line);
            placedRestRects.push(...finalRestRects);
        }
        graphicalStaffEntry.PositionAndShape.calculateBoundingBox();
    }

    private calculateTieCurves(): void {
        for (const musicSystem of this.musicSystems) {
            for (const staffLine of musicSystem.StaffLines) {
                for (const measure of staffLine.Measures) {
                    for (const staffEntry of measure.staffEntries) {
                        for (const graphicalTie of staffEntry.GraphicalTies) {
                            if (graphicalTie.StartNote !== undefined && graphicalTie.StartNote.parentVoiceEntry.parentStaffEntry === staffEntry) {
                                const tieIsAtSystemBreak: boolean = (
                                    graphicalTie.StartNote.parentVoiceEntry.parentStaffEntry.parentMeasure.ParentStaffLine !==
                                    graphicalTie.EndNote.parentVoiceEntry.parentStaffEntry.parentMeasure.ParentStaffLine
                                );
                                this.layoutGraphicalTie(graphicalTie, tieIsAtSystemBreak, measure.ParentStaff.isTab);
                            }
                        }
                    }
                }
            }
        }
    }

    private calculateLyricsPosition(): void {
        const lyricStaffEntriesDict: Dictionary<StaffLine, GraphicalStaffEntry[]> = new Dictionary<StaffLine, GraphicalStaffEntry[]>();
        // sort the lyriceVerseNumbers for every Instrument that has Lyrics
        for (let idx: number = 0, len: number = this.graphicalMusicSheet.ParentMusicSheet.Instruments.length; idx < len; ++idx) {
            const instrument: Instrument = this.graphicalMusicSheet.ParentMusicSheet.Instruments[idx];
            if (instrument.HasLyrics && instrument.LyricVersesNumbers.length > 0) {
                instrument.LyricVersesNumbers.sort();
            }
        }
        // first calc lyrics text positions
        for (let idx2: number = 0, len2: number = this.musicSystems.length; idx2 < len2; ++idx2) {
            const musicSystem: MusicSystem = this.musicSystems[idx2];
            for (let idx3: number = 0, len3: number = musicSystem.StaffLines.length; idx3 < len3; ++idx3) {
                const staffLine: StaffLine = musicSystem.StaffLines[idx3];
                const lyricsStaffEntries: GraphicalStaffEntry[] =
                    this.calculateSingleStaffLineLyricsPosition(staffLine, staffLine.ParentStaff.ParentInstrument.LyricVersesNumbers);
                lyricStaffEntriesDict.setValue(staffLine, lyricsStaffEntries);
                this.calculateLyricsExtendsAndDashes(lyricStaffEntriesDict.getValue(staffLine));
            }
        }
        // then fill in the lyric word dashes and lyrics extends/underscores
        for (let idx2: number = 0, len2: number = this.musicSystems.length; idx2 < len2; ++idx2) {
            const musicSystem: MusicSystem = this.musicSystems[idx2];
            for (let idx3: number = 0, len3: number = musicSystem.StaffLines.length; idx3 < len3; ++idx3) {
                const staffLine: StaffLine = musicSystem.StaffLines[idx3];
                this.calculateLyricsExtendsAndDashes(lyricStaffEntriesDict.getValue(staffLine));
            }
        }
    }

    /**
     * This method calculates the dashes within the syllables of a LyricWord
     * @param lyricEntry
     */
    private calculateSingleLyricWord(lyricEntry: GraphicalLyricEntry): void {
        // const skyBottomLineCalculator: SkyBottomLineCalculator = new SkyBottomLineCalculator (this.rules);
        const graphicalLyricWord: GraphicalLyricWord = lyricEntry.ParentLyricWord;
        const index: number = graphicalLyricWord.GraphicalLyricsEntries.indexOf(lyricEntry);
        let nextLyricEntry: GraphicalLyricEntry = undefined;
        if (index >= 0) {
            nextLyricEntry = graphicalLyricWord.GraphicalLyricsEntries[index + 1];
        }
        if (!nextLyricEntry) {
            return;
        }
        const startStaffLine: StaffLine = <StaffLine>lyricEntry.StaffEntryParent.parentMeasure.ParentStaffLine;
        const nextStaffLine: StaffLine = <StaffLine>nextLyricEntry.StaffEntryParent.parentMeasure.ParentStaffLine;
        const startStaffEntry: GraphicalStaffEntry = lyricEntry.StaffEntryParent;
        const endStaffentry: GraphicalStaffEntry = nextLyricEntry.StaffEntryParent;

        // if on the same StaffLine
        if (lyricEntry.StaffEntryParent.parentMeasure.ParentStaffLine === nextLyricEntry.StaffEntryParent.parentMeasure.ParentStaffLine) {
            // start- and End margins from the text Labels
            const startX: number = startStaffEntry.parentMeasure.PositionAndShape.RelativePosition.x +
                startStaffEntry.PositionAndShape.RelativePosition.x +
                lyricEntry.GraphicalLabel.PositionAndShape.RelativePosition.x +
                lyricEntry.GraphicalLabel.PositionAndShape.BorderMarginRight -
                lyricEntry.GraphicalLabel.CenteringXShift; // TODO not sure why this is necessary, see Christbaum measure 9+11, Land der Berge 11-12

            const endX: number = endStaffentry.parentMeasure.PositionAndShape.RelativePosition.x +
                endStaffentry.PositionAndShape.RelativePosition.x +
                lyricEntry.GraphicalLabel.PositionAndShape.RelativePosition.x +
                nextLyricEntry.GraphicalLabel.PositionAndShape.BorderMarginLeft;
            const y: number = lyricEntry.GraphicalLabel.PositionAndShape.RelativePosition.y;
            let numberOfDashes: number = 1;
            if ((endX - startX) > this.rules.MinimumDistanceBetweenDashes * 3) {
                // *3: need distance between word to first dash, dash to dash, dash to next word
                numberOfDashes = Math.floor((endX - startX) / this.rules.MinimumDistanceBetweenDashes) - 1;
            }
            // check distance and create the adequate number of Dashes
            if (numberOfDashes === 1) {
                // distance between the two GraphicalLyricEntries is big for only one Dash, position in the middle
                this.calculateSingleDashForLyricWord(startStaffLine, startX, endX, y);
            } else {
                // distance is big enough for more Dashes
                // calculate the adequate number of Dashes from the distance between the two LyricEntries
                // distance between the Dashes should be equal
                this.calculateDashes(startStaffLine, startX, endX, y);
            }
        } else {
            // start and end on different StaffLines
            // start margin from the text Label until the End of StaffLine
            const startX: number = startStaffEntry.parentMeasure.PositionAndShape.RelativePosition.x +
                startStaffEntry.PositionAndShape.RelativePosition.x +
                lyricEntry.GraphicalLabel.PositionAndShape.BorderMarginRight;
            const lastGraphicalMeasure: GraphicalMeasure = startStaffLine.Measures[startStaffLine.Measures.length - 1];
            const endX: number = lastGraphicalMeasure.PositionAndShape.RelativePosition.x + lastGraphicalMeasure.PositionAndShape.Size.width;
            let y: number = lyricEntry.GraphicalLabel.PositionAndShape.RelativePosition.y;

            // calculate Dashes for the first StaffLine
            this.calculateDashes(startStaffLine, startX, endX, y);

            // calculate Dashes for the second StaffLine (only if endStaffEntry isn't the first StaffEntry of the StaffLine)
            if (nextStaffLine && // check for undefined objects e.g. when drawingRange given
                nextStaffLine.Measures[0] &&
                endStaffentry.parentMeasure.ParentStaffLine &&
                !(endStaffentry === endStaffentry.parentMeasure.staffEntries[0] &&
                endStaffentry.parentMeasure === endStaffentry.parentMeasure.ParentStaffLine.Measures[0])) {
                const secondStartX: number = nextStaffLine.Measures[0].staffEntries[0].PositionAndShape.RelativePosition.x;
                const secondEndX: number = endStaffentry.parentMeasure.PositionAndShape.RelativePosition.x +
                    endStaffentry.PositionAndShape.RelativePosition.x +
                    nextLyricEntry.GraphicalLabel.PositionAndShape.BorderMarginLeft;
                y = nextLyricEntry.GraphicalLabel.PositionAndShape.RelativePosition.y;
                this.calculateDashes(nextStaffLine, secondStartX, secondEndX, y);
            }
        }
    }

    /**
     * This method calculates Dashes for a LyricWord.
     * @param staffLine
     * @param startX
     * @param endX
     * @param y
     */
    private calculateDashes(staffLine: StaffLine, startX: number, endX: number, y: number): void {
        let distance: number = endX - startX;
        if (distance < this.rules.MinimumDistanceBetweenDashes * 3) {
            this.calculateSingleDashForLyricWord(staffLine, startX, endX, y);
        } else {
            // enough distance for more Dashes
            const numberOfDashes: number = Math.floor(distance / this.rules.MinimumDistanceBetweenDashes) - 1;
            const distanceBetweenDashes: number = distance / (numberOfDashes + 1);
            let counter: number = 0;

            startX += distanceBetweenDashes;
            endX -= distanceBetweenDashes;
            while (counter <= Math.floor(numberOfDashes / 2.0) && endX > startX) {
                distance = this.calculateRightAndLeftDashesForLyricWord(staffLine, startX, endX, y);
                startX += distanceBetweenDashes;
                endX -= distanceBetweenDashes;
                counter++;
            }

            // if the remaining distance isn't big enough for two Dashes,
            // but long enough for a middle dash inbetween,
            // then put the last Dash in the middle of the remaining distance
            if (distance > distanceBetweenDashes * 2) {
                this.calculateSingleDashForLyricWord(staffLine, startX, endX, y);
            }
        }
    }

    /**
     * This method calculates a single Dash for a LyricWord, positioned in the middle of the given distance.
     * @param {StaffLine} staffLine
     * @param {number} startX
     * @param {number} endX
     * @param {number} y
     */
    private calculateSingleDashForLyricWord(staffLine: StaffLine, startX: number, endX: number, y: number): void {
        const label: Label = new Label("-");
        label.colorDefault = this.rules.DefaultColorLyrics; // if undefined, no change. saves an if check
        let textHeight: number = this.rules.LyricsHeight;
        if (endX - startX < 0.8) {
            textHeight *= 0.8;
            y -= 0.1 * textHeight; // dash moves downwards when textHeight is reduced. counteract that.
            //xShift = -0.1;
            // x-position is situational, sometimes it's slightly right-leaning and tends to overlap with the right LyricsEntry
            //   (see Cornelius - Christbaum, measure 9 and 11 ("li-che", "li-ger"), due to centering x-shift = GraphicalLabel.CenteringXShift)
            // sometimes the x-position is perfect and the interval is extremely narrow
            //   (see Mozart/Holzer Land der Berge measure 11-12)
            // or even slightly too far left (Beethoven Geliebte measure 4, due to centering x-shift = GraphicalLabel.CenteringXShift)
        }
        const dash: GraphicalLabel = new GraphicalLabel(
            label, textHeight, TextAlignmentEnum.CenterBottom, this.rules);
        dash.setLabelPositionAndShapeBorders();
        staffLine.LyricsDashes.push(dash);
        if (this.staffLinesWithLyricWords.indexOf(staffLine) === -1) {
            this.staffLinesWithLyricWords.push(staffLine);
        }
        dash.PositionAndShape.Parent = staffLine.PositionAndShape;
        const relative: PointF2D = new PointF2D(startX + (endX - startX) / 2, y);
        dash.PositionAndShape.RelativePosition = relative;
    }

    /**
     * Layouts the underscore line when a lyric entry is marked as extend
     * @param {GraphicalLyricEntry} lyricEntry
     */
    private calculateLyricExtend(lyricEntry: GraphicalLyricEntry): void {
        let startY: number = lyricEntry.GraphicalLabel.PositionAndShape.RelativePosition.y;
        const startStaffEntry: GraphicalStaffEntry = lyricEntry.StaffEntryParent;
        const startStaffLine: StaffLine = startStaffEntry.parentMeasure.ParentStaffLine;

        // find endstaffEntry and staffLine
        let endStaffEntry: GraphicalStaffEntry = undefined;
        let endStaffLine: StaffLine = undefined;
        const staffIndex: number = startStaffEntry.parentMeasure.ParentStaff.idInMusicSheet;
        for (let index: number = startStaffEntry.parentVerticalContainer.Index + 1;
            index < this.graphicalMusicSheet.VerticalGraphicalStaffEntryContainers.length;
            ++index) {
            const gse: GraphicalStaffEntry = this.graphicalMusicSheet.VerticalGraphicalStaffEntryContainers[index].StaffEntries[staffIndex];
            if (!gse) {
                continue;
            }
            if (gse.hasOnlyRests()) {
                break;
            }
            if (gse.LyricsEntries.length > 0) {
                break;
            }
            endStaffEntry = gse;
            endStaffLine = endStaffEntry.parentMeasure.ParentStaffLine;
            if (!endStaffLine) {
                endStaffLine = startStaffEntry.parentMeasure.ParentStaffLine;
            }
        }
        if (!endStaffEntry || !endStaffLine) {
            return;
        }
        // if on the same StaffLine
        if (startStaffLine === endStaffLine && endStaffEntry.parentMeasure.ParentStaffLine) {
            // start- and End margins from the text Labels
            const startX: number = startStaffEntry.parentMeasure.PositionAndShape.RelativePosition.x +
                startStaffEntry.PositionAndShape.RelativePosition.x +
                lyricEntry.GraphicalLabel.PositionAndShape.BorderMarginRight;
            // + startStaffLine.PositionAndShape.AbsolutePosition.x; // doesn't work, done in drawer
            const endX: number = endStaffEntry.parentMeasure.PositionAndShape.RelativePosition.x +
                endStaffEntry.PositionAndShape.RelativePosition.x +
                endStaffEntry.PositionAndShape.BorderMarginRight;
            // + endStaffLine.PositionAndShape.AbsolutePosition.x; // doesn't work, done in drawer
            // TODO maybe add half-width of following note.
            // though we don't have the vexflow note's bbox yet and extend layouting is unconstrained,
            // we have more room for spacing without it.
            // needed in order to line up with the Label's text bottom line (is the y position of the underscore)
            startY -= lyricEntry.GraphicalLabel.PositionAndShape.Size.height / 4;
            // create a Line (as underscore after the LyricLabel's End)
            this.calculateSingleLyricWordWithUnderscore(startStaffLine, startX, endX, startY);
        } else { // start and end on different StaffLines
            // start margin from the text Label until the End of StaffLine
            const lastMeasureBb: BoundingBox = startStaffLine.Measures[startStaffLine.Measures.length - 1].PositionAndShape;
            const startX: number = startStaffEntry.parentMeasure.PositionAndShape.RelativePosition.x +
                startStaffEntry.PositionAndShape.RelativePosition.x +
                lyricEntry.GraphicalLabel.PositionAndShape.BorderMarginRight;
            const endX: number = lastMeasureBb.RelativePosition.x +
                lastMeasureBb.Size.width;
            // needed in order to line up with the Label's text bottom line
            startY -= lyricEntry.GraphicalLabel.PositionAndShape.Size.height / 4;
            // first Underscore until the StaffLine's End
            this.calculateSingleLyricWordWithUnderscore(startStaffLine, startX, endX, startY);
            if (!endStaffEntry) {
                return;
            }
            // second Underscore in the endStaffLine until endStaffEntry (if endStaffEntry isn't the first StaffEntry of the StaffLine))
            if (endStaffEntry.parentMeasure.ParentStaffLine && endStaffEntry.parentMeasure.staffEntries &&
                !(endStaffEntry === endStaffEntry.parentMeasure.staffEntries[0] &&
                endStaffEntry.parentMeasure === endStaffEntry.parentMeasure.ParentStaffLine.Measures[0])) {
                const secondStartX: number = endStaffLine.Measures[0].staffEntries[0].PositionAndShape.RelativePosition.x;
                const secondEndX: number = endStaffEntry.parentMeasure.PositionAndShape.RelativePosition.x +
                    endStaffEntry.PositionAndShape.RelativePosition.x +
                    endStaffEntry.PositionAndShape.BorderMarginRight;
                this.calculateSingleLyricWordWithUnderscore(endStaffLine, secondStartX, secondEndX, startY);
            }
        }
    }

    /**
     * This method calculates a single underscoreLine.
     * @param staffLine
     * @param startX
     * @param end
     * @param y
     */
    private calculateSingleLyricWordWithUnderscore(staffLine: StaffLine, startX: number, endX: number, y: number): void {
        const lineStart: PointF2D = new PointF2D(startX, y);
        const lineEnd: PointF2D = new PointF2D(endX, y);
        const graphicalLine: GraphicalLine = new GraphicalLine(lineStart, lineEnd, this.rules.LyricUnderscoreLineWidth);
        graphicalLine.colorHex = this.rules.DefaultColorLyrics; // if undefined, no change. saves an if check
        staffLine.LyricLines.push(graphicalLine);
        if (this.staffLinesWithLyricWords.indexOf(staffLine) === -1) {
            this.staffLinesWithLyricWords.push(staffLine);
        }
    }

    /**
     * This method calculates two Dashes for a LyricWord, positioned at the the two ends of the given distance.
     * @param {StaffLine} staffLine
     * @param {number} startX
     * @param {number} endX
     * @param {number} y
     * @returns {number}
     */
    private calculateRightAndLeftDashesForLyricWord(staffLine: StaffLine, startX: number, endX: number, y: number): number {
        const leftLabel: Label = new Label("-");
        leftLabel.colorDefault = this.rules.DefaultColorLyrics; // if undefined, no change. saves an if check
        const leftDash: GraphicalLabel = new GraphicalLabel(
            leftLabel, this.rules.LyricsHeight, TextAlignmentEnum.CenterBottom, this.rules);
        leftDash.setLabelPositionAndShapeBorders();
        staffLine.LyricsDashes.push(leftDash);
        if (this.staffLinesWithLyricWords.indexOf(staffLine) === -1) {
            this.staffLinesWithLyricWords.push(staffLine);
        }
        leftDash.PositionAndShape.Parent = staffLine.PositionAndShape;
        const leftDashRelative: PointF2D = new PointF2D(startX, y);
        leftDash.PositionAndShape.RelativePosition = leftDashRelative;

        const rightLabel: Label = new Label("-");
        const rightDash: GraphicalLabel = new GraphicalLabel(
            rightLabel, this.rules.LyricsHeight, TextAlignmentEnum.CenterBottom, this.rules);
        rightDash.setLabelPositionAndShapeBorders();
        staffLine.LyricsDashes.push(rightDash);
        rightDash.PositionAndShape.Parent = staffLine.PositionAndShape;
        const rightDashRelative: PointF2D = new PointF2D(endX, y);
        rightDash.PositionAndShape.RelativePosition = rightDashRelative;
        return (rightDash.PositionAndShape.RelativePosition.x - leftDash.PositionAndShape.RelativePosition.x);
    }

    //So we can track shared notes bounding boxes to avoid collision + skyline issues
    protected dynamicExpressionMap: Map<number, BoundingBox> = new Map<number, BoundingBox>();

    private calculateDynamicExpressions(): void {
        const maxIndex: number = Math.min(this.graphicalMusicSheet.ParentMusicSheet.SourceMeasures.length - 1, this.rules.MaxMeasureToDrawIndex);
        const minIndex: number = Math.min(this.rules.MinMeasureToDrawIndex, this.graphicalMusicSheet.ParentMusicSheet.SourceMeasures.length);
        for (let i: number = minIndex; i <= maxIndex; i++) {
            const sourceMeasure: SourceMeasure = this.graphicalMusicSheet.ParentMusicSheet.SourceMeasures[i];
            //Reset, beginning of new measure
            this.dynamicExpressionMap.clear();
            for (let j: number = 0; j < sourceMeasure.StaffLinkedExpressions.length; j++) {
                if (!this.graphicalMusicSheet.MeasureList[i] || !this.graphicalMusicSheet.MeasureList[i][j]) {
                    continue;
                }

                if (this.graphicalMusicSheet.MeasureList[i][j].ParentStaff.isVisible()) {
                    for (let k: number = 0; k < sourceMeasure.StaffLinkedExpressions[j].length; k++) {
                        if (sourceMeasure.StaffLinkedExpressions[j][k].InstantaneousDynamic !== undefined ||
                            (sourceMeasure.StaffLinkedExpressions[j][k].StartingContinuousDynamic !== undefined &&
                                sourceMeasure.StaffLinkedExpressions[j][k].StartingContinuousDynamic.StartMultiExpression ===
                                sourceMeasure.StaffLinkedExpressions[j][k] && sourceMeasure.StaffLinkedExpressions[j][k].UnknownList.length === 0)
                        ) {
                            this.calculateDynamicExpressionsForMultiExpression(sourceMeasure.StaffLinkedExpressions[j][k], i, j);
                        }
                    }
                }
            }
        }
        this.dynamicExpressionMap.clear();
    }

    private calculateOctaveShifts(): void {
        for (let i: number = 0; i < this.graphicalMusicSheet.ParentMusicSheet.SourceMeasures.length; i++) {
            const sourceMeasure: SourceMeasure = this.graphicalMusicSheet.ParentMusicSheet.SourceMeasures[i];
            for (let j: number = 0; j < sourceMeasure.StaffLinkedExpressions.length; j++) {
                if (!this.graphicalMusicSheet.MeasureList[i] || !this.graphicalMusicSheet.MeasureList[i][j]) {
                    continue;
                }
                if (this.graphicalMusicSheet.MeasureList[i][j].ParentStaff.isVisible()) {
                    for (let k: number = 0; k < sourceMeasure.StaffLinkedExpressions[j].length; k++) {
                        if ((sourceMeasure.StaffLinkedExpressions[j][k].OctaveShiftStart)) {
                            this.calculateSingleOctaveShift(sourceMeasure, sourceMeasure.StaffLinkedExpressions[j][k], i, j);
                        }
                    }
                }
            }
        }
    }

    private calculatePedals(): void {
        for (let i: number = 0; i < this.graphicalMusicSheet.ParentMusicSheet.SourceMeasures.length; i++) {
            const sourceMeasure: SourceMeasure = this.graphicalMusicSheet.ParentMusicSheet.SourceMeasures[i];
            for (let j: number = 0; j < sourceMeasure.StaffLinkedExpressions.length; j++) {
                if (!this.graphicalMusicSheet.MeasureList[i] || !this.graphicalMusicSheet.MeasureList[i][j]) {
                    continue;
                }
                if (this.graphicalMusicSheet.MeasureList[i][j].ParentStaff.isVisible()) {
                    for (let k: number = 0; k < sourceMeasure.StaffLinkedExpressions[j].length; k++) {
                        if ((sourceMeasure.StaffLinkedExpressions[j][k].PedalStart)) {
                            this.calculateSinglePedal(sourceMeasure, sourceMeasure.StaffLinkedExpressions[j][k], i, j);
                        }
                    }
                }
            }
        }
    }

    private calculateWavyLines(): void {
        for (let i: number = 0; i < this.graphicalMusicSheet.ParentMusicSheet.SourceMeasures.length; i++) {
            const sourceMeasure: SourceMeasure = this.graphicalMusicSheet.ParentMusicSheet.SourceMeasures[i];
            for (let j: number = 0; j < sourceMeasure.StaffLinkedExpressions.length; j++) {
                if (!this.graphicalMusicSheet.MeasureList[i] || !this.graphicalMusicSheet.MeasureList[i][j]) {
                    continue;
                }
                if (this.graphicalMusicSheet.MeasureList[i][j].ParentStaff.isVisible()) {
                    for (let k: number = 0; k < sourceMeasure.StaffLinkedExpressions[j].length; k++) {
                        if ((sourceMeasure.StaffLinkedExpressions[j][k].WavyLineStart)) {
                            this.calculateSingleWavyLine(sourceMeasure, sourceMeasure.StaffLinkedExpressions[j][k], i, j);
                        }
                    }
                }
            }
        }
    }

    private getFirstLeftNotNullStaffEntryFromContainer(horizontalIndex: number, verticalIndex: number, multiStaffInstrument: boolean): GraphicalStaffEntry {
        if (this.graphicalMusicSheet.VerticalGraphicalStaffEntryContainers[horizontalIndex].StaffEntries[verticalIndex]) {
            return this.graphicalMusicSheet.VerticalGraphicalStaffEntryContainers[horizontalIndex].StaffEntries[verticalIndex];
        }
        for (let i: number = horizontalIndex - 1; i >= 0; i--) {
            if (this.graphicalMusicSheet.VerticalGraphicalStaffEntryContainers[i].StaffEntries[verticalIndex]) {
                return this.graphicalMusicSheet.VerticalGraphicalStaffEntryContainers[i].StaffEntries[verticalIndex];
            }
        }
        return undefined;
    }

    private getFirstRightNotNullStaffEntryFromContainer(horizontalIndex: number, verticalIndex: number, multiStaffInstrument: boolean): GraphicalStaffEntry {
        if (this.graphicalMusicSheet.VerticalGraphicalStaffEntryContainers[horizontalIndex].StaffEntries[verticalIndex]) {
            return this.graphicalMusicSheet.VerticalGraphicalStaffEntryContainers[horizontalIndex].StaffEntries[verticalIndex];
        }
        for (let i: number = horizontalIndex + 1; i < this.graphicalMusicSheet.VerticalGraphicalStaffEntryContainers.length; i++) {
            if (this.graphicalMusicSheet.VerticalGraphicalStaffEntryContainers[i].StaffEntries[verticalIndex]) {
                return this.graphicalMusicSheet.VerticalGraphicalStaffEntryContainers[i].StaffEntries[verticalIndex];
            }
        }
        return undefined;
    }

    private calculateWordRepetitionInstructions(): void {
        // Track currently active volta spans (can have multiple nested or sequential)
        // Each span tracks: startMeasure index, endingIndices
        const activeVoltaSpans: {startMeasure: number, endingIndices: number[]}[] = [];

        for (let i: number = 0; i < this.graphicalMusicSheet.ParentMusicSheet.SourceMeasures.length; i++) {
            const sourceMeasure: SourceMeasure = this.graphicalMusicSheet.ParentMusicSheet.SourceMeasures[i];

            // Check if this measure has a Begin or End ending instruction
            let hasBeginEnding: boolean = false;
            let hasEndEnding: boolean = false;
            let beginEndingIndices: number[] = undefined;

            for (const instruction of sourceMeasure.FirstRepetitionInstructions) {
                if (instruction.type === RepetitionInstructionEnum.Ending && instruction.alignment === AlignmentType.Begin) {
                    hasBeginEnding = true;
                    beginEndingIndices = instruction.endingIndices;
                }
            }

            for (const instruction of sourceMeasure.LastRepetitionInstructions) {
                if (instruction.type === RepetitionInstructionEnum.Ending &&
                    (instruction.alignment === AlignmentType.End || instruction.alignment === AlignmentType.Discontinue)) {
                    hasEndEnding = true;
                }
            }

            // If this is a Begin ending, start tracking a new volta span
            if (hasBeginEnding) {
                activeVoltaSpans.push({startMeasure: i, endingIndices: beginEndingIndices});
            }

            // Process all regular instructions
            for (let idx: number = 0, len: number = sourceMeasure.FirstRepetitionInstructions.length; idx < len; ++idx) {
                const instruction: RepetitionInstruction = sourceMeasure.FirstRepetitionInstructions[idx];
                this.calculateWordRepetitionInstruction(instruction, i);
            }
            for (let idx: number = 0, len: number = sourceMeasure.LastRepetitionInstructions.length; idx < len; ++idx) {
                const instruction: RepetitionInstruction = sourceMeasure.LastRepetitionInstructions[idx];
                this.calculateWordRepetitionInstruction(instruction, i);
            }

            // Add continuing volta line:
            //   If this measure has no Begin or End volta/instruction but we have active volta spans,
            //   and this measure is AFTER the start of the span (not the same measure),
            //   then add a MID volta. (continuing volta line in a measure between start and end volta measure)
            if (!hasBeginEnding && !hasEndEnding && activeVoltaSpans.length > 0) {
                // Use the most recent active span
                const activeSpan: {startMeasure: number, endingIndices: number[]} = activeVoltaSpans[activeVoltaSpans.length - 1];
                if (i > activeSpan.startMeasure) {
                    const midInstruction: RepetitionInstruction = new RepetitionInstruction(
                        i, RepetitionInstructionEnum.Ending, AlignmentType.Mid, undefined, activeSpan.endingIndices
                    );
                    this.calculateWordRepetitionInstruction(midInstruction, i);
                }
            }

            // If this is an End ending, close the active volta span
            if (hasEndEnding && activeVoltaSpans.length > 0) {
                activeVoltaSpans.pop();
            }
        }
    }

    private calculateRepetitionEndings(): void {
        const musicsheet: MusicSheet = this.graphicalMusicSheet.ParentMusicSheet;
        for (let idx: number = 0, len: number = musicsheet.Repetitions.length; idx < len; ++idx) {
            const repetition: Repetition = musicsheet.Repetitions[idx];
            this.calcGraphicalRepetitionEndingsRecursively(repetition);
        }
    }

    private calculateTempoExpressions(): void {
        const maxIndex: number = Math.min(this.graphicalMusicSheet.ParentMusicSheet.SourceMeasures.length - 1, this.rules.MaxMeasureToDrawIndex);
        const minIndex: number = this.rules.MinMeasureToDrawIndex;
        for (let i: number = minIndex; i <= maxIndex; i++) {
            const sourceMeasure: SourceMeasure = this.graphicalMusicSheet.ParentMusicSheet.SourceMeasures[i];
            for (let j: number = 0; j < sourceMeasure.TempoExpressions.length; j++) {
                this.calculateTempoExpressionsForMultiTempoExpression(sourceMeasure, sourceMeasure.TempoExpressions[j], i);
            }
        }
    }

    private calculateRehearsalMarks(): void {
        if (!this.rules.RenderRehearsalMarks) {
            return;
        }
        for (const measure of this.graphicalMusicSheet.ParentMusicSheet.SourceMeasures) {
            this.calculateRehearsalMark(measure);
        }
    }

    protected calculateRehearsalMark(measure: SourceMeasure): void {
        throw new Error(this.abstractNotImplementedErrorMessage);
    }

    private calculateMoodAndUnknownExpressions(): void {
        for (let i: number = 0; i < this.graphicalMusicSheet.ParentMusicSheet.SourceMeasures.length; i++) {
            const sourceMeasure: SourceMeasure = this.graphicalMusicSheet.ParentMusicSheet.SourceMeasures[i];
            for (let j: number = 0; j < sourceMeasure.StaffLinkedExpressions.length; j++) {
                if (!this.graphicalMusicSheet.MeasureList[i] || !this.graphicalMusicSheet.MeasureList[i][j]) {
                    continue;
                }
                let staffIndexToUse: number = j;
                if (!this.graphicalMusicSheet.MeasureList[i][j].ParentStaff.isVisible()) {
                    // If this staff is hidden and it's the first staff (j === 0) (see #1621),
                    // find the first visible staff to render expressions on (similar to rehearsal mark fix #1555)
                    if (j === 0) {
                        let foundVisibleStaff: boolean = false;
                        for (let s: number = 0; s < this.graphicalMusicSheet.MeasureList[i].length; s++) {
                            if (this.graphicalMusicSheet.MeasureList[i][s]?.ParentStaff.isVisible()) {
                                staffIndexToUse = s; // render words on this staff instead of first/invisible staff (see #1621)
                                foundVisibleStaff = true;
                                break;
                            }
                        }
                        if (!foundVisibleStaff) {
                            continue; // No visible staff found, skip this expression
                        }
                    } else {
                        continue; // Non-first staff is hidden, skip its expressions
                    }
                }
                for (let k: number = 0; k < sourceMeasure.StaffLinkedExpressions[j].length; k++) {
                    if ((sourceMeasure.StaffLinkedExpressions[j][k].MoodList.length > 0) ||
                        (sourceMeasure.StaffLinkedExpressions[j][k].UnknownList.length > 0)) {
                        this.calculateMoodAndUnknownExpression(sourceMeasure.StaffLinkedExpressions[j][k], i, staffIndexToUse);
                    }
                }
            }
        }
    }

    /**
     * Calculates the desired stem direction depending on the number (or type) of voices.
     * If more than one voice is there, the main voice (typically the first or upper voice) will get stem up direction.
     * The others get stem down direction.
     * @param voiceEntry the voiceEntry for which the stem direction has to be calculated
     */
    private calculateStemDirectionFromVoices(voiceEntry: VoiceEntry): void {
        // Stem direction calculation:
        const hasLink: boolean = voiceEntry.ParentSourceStaffEntry.Link !== undefined;
        if (hasLink) {
            // in case of StaffEntryLink don't check mainVoice / linkedVoice
            if (voiceEntry === voiceEntry.ParentSourceStaffEntry.VoiceEntries[0]) {
                // set stem up:
                voiceEntry.WantedStemDirection = StemDirectionType.Up;
                return;
            } else {
                // set stem down:
                voiceEntry.WantedStemDirection = StemDirectionType.Down;
                return;
            }
        } else {
            if (voiceEntry.ParentVoice instanceof LinkedVoice) {
                // Linked voice: set stem down:
                voiceEntry.WantedStemDirection = StemDirectionType.Down;
            } else {
                // if this voiceEntry belongs to the mainVoice:
                // check first that there are also more voices present:
                if (voiceEntry.ParentSourceStaffEntry.VoiceEntries.length > 1) {
                    // as this voiceEntry belongs to the mainVoice: stem Up
                    voiceEntry.WantedStemDirection = StemDirectionType.Up;
                }
            }
        }
        // setBeamNotesWantedStemDirections() will be called at end of measure (createGraphicalMeasure)
    }

    /** Sets a voiceEntry's stem direction to one already set in other notes in its beam, if it has one. */
    private setBeamNotesWantedStemDirections(voiceEntry: VoiceEntry): void {
        if (!(voiceEntry.Notes.length > 0)) {
            return;
        }
        // don't just set direction if undefined. if there's a note in the beam with a different stem direction, Vexflow draws it with an unending stem.
        // if (voiceEntry.WantedStemDirection === StemDirectionType.Undefined) {
        const beam: Beam = voiceEntry.Notes[0].NoteBeam;
        if (beam) {
            // if there is a beam, find any already set stemDirection in the beam:
            // Skip hidden notes (notehead="none") when determining beam direction
            for (const note of beam.Notes) {
                // if (note.ParentVoiceEntry === voiceEntry) {
                //     continue; // this could cause a misreading, also potentially in cross-staf beams, in any case it's unnecessary.
                //} else if

                // Skip notes with NONE notehead - they shouldn't influence beam direction
                if (note.Notehead?.Shape === NoteHeadShape.NONE) {
                    continue;
                }

                if (note.ParentVoiceEntry.WantedStemDirection !== StemDirectionType.Undefined) {
                    if (note.ParentVoiceEntry.ParentSourceStaffEntry.ParentStaff.Id === voiceEntry.ParentSourceStaffEntry.ParentStaff.Id) {
                        // set the stem direction
                        voiceEntry.WantedStemDirection = note.ParentVoiceEntry.WantedStemDirection;
                        break;
                    }
                }
            }
        }
    }
}
