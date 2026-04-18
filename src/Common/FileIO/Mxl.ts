import { IXmlElement } from "./Xml";
import JSZip from "jszip";
import log from "loglevel";

export class MXLFile {
    private blob: Blob;
    public zipFile: JSZip;
    public xmlText: string;
    /** Set after tryUnzip(). True if it could be unzipped successfully. */
    public unzipSuccessful: boolean = false;
    public constructor(blob: Blob) {
        this.blob = blob;
    }

    /** Try unzipping to see if this is a zip file.
     * This is a separate method so that we don't need to unzip twice to check whether it's a zip file.
     */
    public async tryUnzip(): Promise<boolean> {
        this.zipFile = new JSZip();
        try {
            this.unzipSuccessful = true;
            await this.zipFile.loadAsync(this.blob);
            return true;
        } catch (e) {
            this.unzipSuccessful = false;
            return false;
        }
    }

    public getXmlString(): Promise<string> {
        return MXLHelper.jszipToXMLstring(this.zipFile);
    }
}

/**
 * Some helper methods to handle MXL files.
 */
export class MXLHelper {
    /** Returns the documentElement of MXL data. */
    public static MXLtoIXmlElement(data: string): Promise<IXmlElement> {
        return this.MXLtoXMLstring(data)
        .then(
            (content: string) => {
                const parser: DOMParser = new DOMParser();
                const xml: Document = parser.parseFromString(content, "text/xml");
                const doc: IXmlElement = new IXmlElement(xml.documentElement);
                return Promise.resolve(doc);
            },
            (err: any) => {
                throw new Error("extractSheetFromMxl: " + err.message);
            }
        );
    }

    public static async jszipToXMLstring(zip: JSZip): Promise<string> {
        const entryNames: string[] = Object.keys(zip.files).filter((path: string) => !zip.files[path].dir);
        const candidatePaths: string[] = [];
        const pushCandidate: (path: string) => void = (path: string): void => {
            if (path && zip.file(path) && !candidatePaths.includes(path)) {
                candidatePaths.push(path);
            }
        };

        const containerPath: string = "META-INF/container.xml";
        if (zip.file(containerPath)) {
            const container: string = await this.readXmlTextFromZipPath(zip, containerPath);
            const parser: DOMParser = new DOMParser();
            const doc: Document = parser.parseFromString(container, "text/xml");
            const rootFiles: HTMLCollectionOf<Element> = doc.getElementsByTagName("rootfile");
            for (let i: number = 0; i < rootFiles.length; i += 1) {
                const rootFile: string = rootFiles[i].getAttribute("full-path");
                pushCandidate(rootFile);
            }
        }

        const extensionsInPriority: RegExp[] = [/\.musicxml$/i, /\.xml$/i, /\.mscx$/i];
        for (const extension of extensionsInPriority) {
            for (const path of entryNames) {
                if (extension.test(path)) {
                    pushCandidate(path);
                }
            }
        }

        let foundMuseScoreMscx: boolean = false;
        for (const candidatePath of candidatePaths) {
            const xmlText: string = await this.readXmlTextFromZipPath(zip, candidatePath);
            if (this.isScorePartwiseXml(xmlText)) {
                return xmlText;
            }
            if (candidatePath.toLowerCase().endsWith(".mscx")) {
                foundMuseScoreMscx = true;
            }
        }

        if (foundMuseScoreMscx) {
            throw new Error("Archive contains a MuseScore .mscx score, but no MusicXML score-partwise file.");
        }
        throw new Error("Could not find a MusicXML score-partwise file in archive.");
    }

    private static async readXmlTextFromZipPath(zip: JSZip, filePath: string): Promise<string> {
        const zipObject: JSZip.JSZipObject = zip.file(filePath);
        if (!zipObject) {
            throw new Error("Could not read archive entry: " + filePath);
        }
        let xmlText: string = await zipObject.async("text");
        if (!this.startsWithXmlTag(xmlText)) {
            const uint8Array: Uint8Array = await zipObject.async("uint8array");
            xmlText = new TextDecoder("utf-8").decode(uint8Array);
        }
        if (!this.startsWithXmlTag(xmlText)) {
            const uint8Array: Uint8Array = await zipObject.async("uint8array");
            xmlText = new TextDecoder("utf-16").decode(uint8Array);
        }
        return xmlText;
    }

    private static startsWithXmlTag(content: string): boolean {
        return /^\s*</.test(content);
    }

    private static isScorePartwiseXml(xmlText: string): boolean {
        if (!this.startsWithXmlTag(xmlText)) {
            return false;
        }
        const parser: DOMParser = new DOMParser();
        const xml: Document = parser.parseFromString(xmlText, "text/xml");
        return xml.documentElement?.nodeName?.toLowerCase() === "score-partwise";
    }

    public static MXLtoXMLstring(data: string | Blob): Promise<string> {
        const zip:  JSZip = new JSZip();
        return zip.loadAsync(data).then(
            async (_: any) => {
                return this.jszipToXMLstring(zip);
            },
            (err: any) => {
                log.error(err);
                throw err;
            }
        );
    }
}
