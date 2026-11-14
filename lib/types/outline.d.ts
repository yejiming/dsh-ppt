import { z } from 'zod';
import type { ArtDirection } from './art-direction.ts';
import { type DiscoveredFont } from './fonts.ts';
import { type PptTheme } from './themes.ts';
export { SLIDE_LAYOUTS, SLIDE_TYPES } from './slide-taxonomy.ts';
export declare const OutlineContentItemSchema: z.ZodDiscriminatedUnion<[z.ZodObject<{
    kind: z.ZodLiteral<"point">;
    text: z.ZodPipe<z.ZodPipe<z.ZodString, z.ZodTransform<string, string>>, z.ZodString>;
    label: z.ZodOptional<z.ZodPipe<z.ZodPipe<z.ZodString, z.ZodTransform<string, string>>, z.ZodString>>;
    group: z.ZodOptional<z.ZodPipe<z.ZodPipe<z.ZodString, z.ZodTransform<string, string>>, z.ZodString>>;
    level: z.ZodDefault<z.ZodUnion<readonly [z.ZodLiteral<1>, z.ZodLiteral<2>]>>;
    emphasis: z.ZodDefault<z.ZodBoolean>;
}, z.core.$strict>, z.ZodObject<{
    kind: z.ZodLiteral<"data">;
    label: z.ZodPipe<z.ZodPipe<z.ZodString, z.ZodTransform<string, string>>, z.ZodString>;
    value: z.ZodUnion<readonly [z.ZodNumber, z.ZodPipe<z.ZodPipe<z.ZodString, z.ZodTransform<string, string>>, z.ZodString>]>;
    unit: z.ZodOptional<z.ZodPipe<z.ZodPipe<z.ZodString, z.ZodTransform<string, string>>, z.ZodString>>;
    source: z.ZodOptional<z.ZodPipe<z.ZodPipe<z.ZodString, z.ZodTransform<string, string>>, z.ZodString>>;
    note: z.ZodOptional<z.ZodPipe<z.ZodPipe<z.ZodString, z.ZodTransform<string, string>>, z.ZodString>>;
    group: z.ZodOptional<z.ZodPipe<z.ZodPipe<z.ZodString, z.ZodTransform<string, string>>, z.ZodString>>;
    emphasis: z.ZodDefault<z.ZodBoolean>;
}, z.core.$strict>, z.ZodObject<{
    kind: z.ZodLiteral<"image">;
    role: z.ZodEnum<{
        hero: "hero";
        diagram: "diagram";
        supporting: "supporting";
        background: "background";
        portrait: "portrait";
        logo: "logo";
    }>;
    intent: z.ZodPipe<z.ZodPipe<z.ZodString, z.ZodTransform<string, string>>, z.ZodString>;
    query: z.ZodOptional<z.ZodPipe<z.ZodPipe<z.ZodString, z.ZodTransform<string, string>>, z.ZodString>>;
    asset: z.ZodOptional<z.ZodPipe<z.ZodPipe<z.ZodString, z.ZodTransform<string, string>>, z.ZodString>>;
    caption: z.ZodOptional<z.ZodPipe<z.ZodPipe<z.ZodString, z.ZodTransform<string, string>>, z.ZodString>>;
    group: z.ZodOptional<z.ZodPipe<z.ZodPipe<z.ZodString, z.ZodTransform<string, string>>, z.ZodString>>;
}, z.core.$strict>, z.ZodObject<{
    kind: z.ZodLiteral<"chart">;
    chart_type: z.ZodEnum<{
        table: "table";
        area: "area";
        line: "line";
        waterfall: "waterfall";
        bar: "bar";
        pie: "pie";
        donut: "donut";
        scatter: "scatter";
        bubble: "bubble";
        radar: "radar";
        funnel: "funnel";
        heatmap: "heatmap";
    }>;
    subject: z.ZodPipe<z.ZodPipe<z.ZodString, z.ZodTransform<string, string>>, z.ZodString>;
    data_ref: z.ZodOptional<z.ZodPipe<z.ZodPipe<z.ZodString, z.ZodTransform<string, string>>, z.ZodString>>;
    takeaway: z.ZodPipe<z.ZodPipe<z.ZodString, z.ZodTransform<string, string>>, z.ZodString>;
    group: z.ZodOptional<z.ZodPipe<z.ZodPipe<z.ZodString, z.ZodTransform<string, string>>, z.ZodString>>;
}, z.core.$strict>, z.ZodObject<{
    kind: z.ZodLiteral<"note">;
    purpose: z.ZodEnum<{
        speaker: "speaker";
        production: "production";
    }>;
    text: z.ZodPipe<z.ZodPipe<z.ZodString, z.ZodTransform<string, string>>, z.ZodString>;
}, z.core.$strict>], "kind">;
export declare const PptOutlineSchema: z.ZodArray<z.ZodObject<{
    page: z.ZodNumber;
    type: z.ZodEnum<{
        process: "process";
        quote: "quote";
        data: "data";
        content: "content";
        section: "section";
        summary: "summary";
        cover: "cover";
        agenda: "agenda";
        comparison: "comparison";
        timeline: "timeline";
        ending: "ending";
    }>;
    title: z.ZodPipe<z.ZodPipe<z.ZodString, z.ZodTransform<string, string>>, z.ZodString>;
    content: z.ZodArray<z.ZodDiscriminatedUnion<[z.ZodObject<{
        kind: z.ZodLiteral<"point">;
        text: z.ZodPipe<z.ZodPipe<z.ZodString, z.ZodTransform<string, string>>, z.ZodString>;
        label: z.ZodOptional<z.ZodPipe<z.ZodPipe<z.ZodString, z.ZodTransform<string, string>>, z.ZodString>>;
        group: z.ZodOptional<z.ZodPipe<z.ZodPipe<z.ZodString, z.ZodTransform<string, string>>, z.ZodString>>;
        level: z.ZodDefault<z.ZodUnion<readonly [z.ZodLiteral<1>, z.ZodLiteral<2>]>>;
        emphasis: z.ZodDefault<z.ZodBoolean>;
    }, z.core.$strict>, z.ZodObject<{
        kind: z.ZodLiteral<"data">;
        label: z.ZodPipe<z.ZodPipe<z.ZodString, z.ZodTransform<string, string>>, z.ZodString>;
        value: z.ZodUnion<readonly [z.ZodNumber, z.ZodPipe<z.ZodPipe<z.ZodString, z.ZodTransform<string, string>>, z.ZodString>]>;
        unit: z.ZodOptional<z.ZodPipe<z.ZodPipe<z.ZodString, z.ZodTransform<string, string>>, z.ZodString>>;
        source: z.ZodOptional<z.ZodPipe<z.ZodPipe<z.ZodString, z.ZodTransform<string, string>>, z.ZodString>>;
        note: z.ZodOptional<z.ZodPipe<z.ZodPipe<z.ZodString, z.ZodTransform<string, string>>, z.ZodString>>;
        group: z.ZodOptional<z.ZodPipe<z.ZodPipe<z.ZodString, z.ZodTransform<string, string>>, z.ZodString>>;
        emphasis: z.ZodDefault<z.ZodBoolean>;
    }, z.core.$strict>, z.ZodObject<{
        kind: z.ZodLiteral<"image">;
        role: z.ZodEnum<{
            hero: "hero";
            diagram: "diagram";
            supporting: "supporting";
            background: "background";
            portrait: "portrait";
            logo: "logo";
        }>;
        intent: z.ZodPipe<z.ZodPipe<z.ZodString, z.ZodTransform<string, string>>, z.ZodString>;
        query: z.ZodOptional<z.ZodPipe<z.ZodPipe<z.ZodString, z.ZodTransform<string, string>>, z.ZodString>>;
        asset: z.ZodOptional<z.ZodPipe<z.ZodPipe<z.ZodString, z.ZodTransform<string, string>>, z.ZodString>>;
        caption: z.ZodOptional<z.ZodPipe<z.ZodPipe<z.ZodString, z.ZodTransform<string, string>>, z.ZodString>>;
        group: z.ZodOptional<z.ZodPipe<z.ZodPipe<z.ZodString, z.ZodTransform<string, string>>, z.ZodString>>;
    }, z.core.$strict>, z.ZodObject<{
        kind: z.ZodLiteral<"chart">;
        chart_type: z.ZodEnum<{
            table: "table";
            area: "area";
            line: "line";
            waterfall: "waterfall";
            bar: "bar";
            pie: "pie";
            donut: "donut";
            scatter: "scatter";
            bubble: "bubble";
            radar: "radar";
            funnel: "funnel";
            heatmap: "heatmap";
        }>;
        subject: z.ZodPipe<z.ZodPipe<z.ZodString, z.ZodTransform<string, string>>, z.ZodString>;
        data_ref: z.ZodOptional<z.ZodPipe<z.ZodPipe<z.ZodString, z.ZodTransform<string, string>>, z.ZodString>>;
        takeaway: z.ZodPipe<z.ZodPipe<z.ZodString, z.ZodTransform<string, string>>, z.ZodString>;
        group: z.ZodOptional<z.ZodPipe<z.ZodPipe<z.ZodString, z.ZodTransform<string, string>>, z.ZodString>>;
    }, z.core.$strict>, z.ZodObject<{
        kind: z.ZodLiteral<"note">;
        purpose: z.ZodEnum<{
            speaker: "speaker";
            production: "production";
        }>;
        text: z.ZodPipe<z.ZodPipe<z.ZodString, z.ZodTransform<string, string>>, z.ZodString>;
    }, z.core.$strict>], "kind">>;
    style: z.ZodObject<{
        layout: z.ZodEnum<{
            split: "split";
            "full-bleed": "full-bleed";
            closing: "closing";
            center: "center";
            cover: "cover";
            "title-content": "title-content";
            "two-column": "two-column";
            "three-column": "three-column";
            grid: "grid";
            "hero-image": "hero-image";
            "image-left": "image-left";
            "image-right": "image-right";
            "timeline-horizontal": "timeline-horizontal";
            "timeline-vertical": "timeline-vertical";
            "process-horizontal": "process-horizontal";
            "process-vertical": "process-vertical";
            "chart-focus": "chart-focus";
            "quote-focus": "quote-focus";
        }>;
        background: z.ZodEnum<{
            accent: "accent";
            image: "image";
            light: "light";
            dark: "dark";
        }>;
        accent: z.ZodPipe<z.ZodString, z.ZodTransform<string, string>>;
        title_font: z.ZodPipe<z.ZodPipe<z.ZodString, z.ZodTransform<string, string>>, z.ZodString>;
        body_font: z.ZodPipe<z.ZodPipe<z.ZodString, z.ZodTransform<string, string>>, z.ZodString>;
        visual_direction: z.ZodPipe<z.ZodPipe<z.ZodString, z.ZodTransform<string, string>>, z.ZodString>;
    }, z.core.$strict>;
}, z.core.$strict>>;
export type PptOutline = z.infer<typeof PptOutlineSchema>;
export interface OutlineWriteResult {
    artifact_dir: string;
    outline_path: string;
    design_plan_path?: string;
    design_status: 'directed' | 'legacy';
    page_count: number;
    type_counts: Record<string, number>;
    fonts: string[];
    warnings: string[];
    blocking_warnings: string[];
    /** Present only when the caller pinned the deck to a built-in theme. */
    theme?: {
        id: string;
        name: string;
        palette_source: string;
        accent: string;
        accent_inverted: string;
        findings: string[];
    };
}
/**
 * Check that an authored outline and plan still belong to the theme they name.
 * Selecting a theme and then using unrelated colours or fonts is the failure
 * this catches: the deck reads as neither the theme nor the brief.
 */
export declare function themeConformanceFindings(theme: PptTheme, outline: PptOutline, designPlan: ArtDirection | undefined): string[];
export declare function resolveTheme(themeId: string): PptTheme;
export interface OutlineFontResolutionOptions {
    discovered: readonly DiscoveredFont[];
    platform?: NodeJS.Platform;
}
export declare function validatePptOutline(value: unknown): PptOutline;
export declare function writePptOutline(workspace: string, artifactTitle: string, value: unknown, outputRoot?: string, signal?: AbortSignal, artDirection?: unknown, fontResolution?: OutlineFontResolutionOptions, themeId?: string): Promise<OutlineWriteResult>;
