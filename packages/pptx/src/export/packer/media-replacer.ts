import type { IMediaData } from "@file/media/data";
import type { Media } from "@file/media/media";

export class MediaReplacer {
    public replace(xmlData: string, mediaData: readonly IMediaData[], offset: number): string {
        let currentXmlData = xmlData;
        mediaData.forEach((media, i) => {
            currentXmlData = currentXmlData.replace(
                new RegExp(`\\{media:${this.escapeRegex(media.fileName)}\\}`, "g"),
                (offset + i).toString(),
            );
        });
        return currentXmlData;
    }

    public getMediaRefs(xmlData: string, media: Media): readonly IMediaData[] {
        const mediaPlaceholders = xmlData.matchAll(/\{media:([^}]+)\}/g);
        const keys = new Set<string>();
        for (const match of mediaPlaceholders) {
            keys.add(match[1]);
        }
        return media.Array.filter((m) => keys.has(m.fileName));
    }

    private escapeRegex(str: string): string {
        return str.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    }
}
