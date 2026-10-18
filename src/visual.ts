import powerbi from "powerbi-visuals-api";
import IVisual = powerbi.extensibility.visual.IVisual;
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import EnumerateVisualObjectInstancesOptions = powerbi.EnumerateVisualObjectInstancesOptions;
import VisualObjectInstance = powerbi.VisualObjectInstance;
import DataViewObjects = powerbi.DataViewObjects;

interface Settings {
    orientation: string; fontFamily: string; fontSize: number; bold: boolean;
    fontColor: string; hAlign: string; vAlign: string; autoFit: boolean;
    padding: number; backgroundColor: string; transparency: number;
    showBorder: boolean; borderColor: string; borderWidth: number; cornerRadius: number;
}

const DEFAULTS: Settings = {
    orientation: "bottomUp", fontFamily: "Segoe UI", fontSize: 22, bold: false,
    fontColor: "#252525", hAlign: "center", vAlign: "center", autoFit: true,
    padding: 6, backgroundColor: "#ffffff", transparency: 100,
    showBorder: false, borderColor: "#777777", borderWidth: 1, cornerRadius: 0
};

function property(objects: DataViewObjects | undefined, objectName: string, propertyName: string): any {
    const object: any = objects && (objects as any)[objectName];
    return object && object[propertyName] !== undefined ? object[propertyName] : undefined;
}
function stringProp(objects: DataViewObjects | undefined, group: string, key: string, fallback: string): string {
    const raw = property(objects, group, key);
    return typeof raw === "string" ? raw : fallback;
}
function numberProp(objects: DataViewObjects | undefined, group: string, key: string, fallback: number, min: number, max: number): number {
    const raw = Number(property(objects, group, key));
    return Number.isFinite(raw) ? Math.min(max, Math.max(min, raw)) : fallback;
}
function boolProp(objects: DataViewObjects | undefined, group: string, key: string, fallback: boolean): boolean {
    const raw = property(objects, group, key);
    return typeof raw === "boolean" ? raw : fallback;
}
function colorProp(objects: DataViewObjects | undefined, group: string, key: string, fallback: string): string {
    const raw = property(objects, group, key);
    return raw && raw.solid && typeof raw.solid.color === "string" ? raw.solid.color : fallback;
}
function settingsFrom(objects: DataViewObjects | undefined): Settings {
    return {
        orientation: stringProp(objects, "direction", "orientation", DEFAULTS.orientation),
        fontFamily: stringProp(objects, "textStyle", "fontFamily", DEFAULTS.fontFamily),
        fontSize: numberProp(objects, "textStyle", "fontSize", DEFAULTS.fontSize, 6, 300),
        bold: boolProp(objects, "textStyle", "bold", DEFAULTS.bold),
        fontColor: colorProp(objects, "textStyle", "fontColor", DEFAULTS.fontColor),
        hAlign: stringProp(objects, "textStyle", "hAlign", DEFAULTS.hAlign),
        vAlign: stringProp(objects, "textStyle", "vAlign", DEFAULTS.vAlign),
        autoFit: boolProp(objects, "textStyle", "autoFit", DEFAULTS.autoFit),
        padding: numberProp(objects, "textStyle", "padding", DEFAULTS.padding, 0, 100),
        backgroundColor: colorProp(objects, "cardBackground", "backgroundColor", DEFAULTS.backgroundColor),
        transparency: numberProp(objects, "cardBackground", "transparency", DEFAULTS.transparency, 0, 100),
        showBorder: boolProp(objects, "cardBackground", "showBorder", DEFAULTS.showBorder),
        borderColor: colorProp(objects, "cardBackground", "borderColor", DEFAULTS.borderColor),
        borderWidth: numberProp(objects, "cardBackground", "borderWidth", DEFAULTS.borderWidth, 0, 20),
        cornerRadius: numberProp(objects, "cardBackground", "cornerRadius", DEFAULTS.cornerRadius, 0, 100)
    };
}
function fill(color: string): any { return { solid: { color } }; }

export class Visual implements IVisual {
    private root: HTMLElement;
    private svg: SVGSVGElement;
    private background: SVGRectElement;
    private label: SVGTextElement;
    private settings: Settings = { ...DEFAULTS };
    private value: string = "";
    private width: number = 0;
    private height: number = 0;

    constructor(options?: VisualConstructorOptions) {
        this.root = document.createElement("div");
        this.root.className = "verticalMeasureCardRoot";
        this.root.style.cssText = "width:100%;height:100%;overflow:hidden;box-sizing:border-box";
        this.svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
        this.svg.setAttribute("class", "verticalMeasureCardSvg");
        this.svg.style.cssText = "width:100%;height:100%;display:block;overflow:hidden";
        this.background = document.createElementNS("http://www.w3.org/2000/svg", "rect");
        this.label = document.createElementNS("http://www.w3.org/2000/svg", "text");
        this.label.setAttribute("xml:space", "preserve");
        this.svg.appendChild(this.background);
        this.svg.appendChild(this.label);
        this.root.appendChild(this.svg);
        if (options && options.element) options.element.appendChild(this.root);
    }

    public update(options: VisualUpdateOptions): void {
        this.width = Math.max(0, options.viewport.width);
        this.height = Math.max(0, options.viewport.height);
        const view = options.dataViews && options.dataViews[0];
        this.settings = settingsFrom(view && view.metadata && view.metadata.objects);
        const single = view && view.single;
        const raw: any = single ? single.value : null;
        const column = view && view.metadata && view.metadata.columns && view.metadata.columns[0];
        // Power BI can deliver formattedValue on single values; otherwise host formatting is
        // reconstructed from the column format string without forcing a specific number style.
        const formatted: any = single && (single as any).formattedValue;
        this.value = raw === null || raw === undefined ? "" :
            typeof formatted === "string" ? formatted : this.formatValue(raw, column && column.format);
        this.render();
    }

    private formatValue(value: any, format?: string): string {
        if (typeof value !== "number" || !Number.isFinite(value)) return String(value);
        const f = (format || "").split(";")[0];
        const percent = f.includes("%");
        const currencyToken = f.match(/(?:\$|€|£|¥|₹|₽|﷼)/);
        const decimalsMatch = f.match(/\.(0+|#+|0+#*)/);
        const decimals = decimalsMatch ? decimalsMatch[1].replace(/#/g, "").length : (f ? 0 : 2);
        const useGrouping = f.includes(",");
        const fixedDecimals = Math.min(decimals, 12);
        const num = percent ? value * 100 : value;
        const formatted = new Intl.NumberFormat(undefined, {
            minimumFractionDigits: fixedDecimals, maximumFractionDigits: fixedDecimals,
            useGrouping
        }).format(num);
        return (currencyToken ? currencyToken[0] : "") + formatted + (percent ? "%" : "");
    }

    private render(): void {
        const s = this.settings, w = this.width, h = this.height;
        this.svg.setAttribute("viewBox", `0 0 ${w} ${h}`);
        this.background.setAttribute("width", String(w));
        this.background.setAttribute("height", String(h));
        this.background.setAttribute("rx", String(s.cornerRadius));
