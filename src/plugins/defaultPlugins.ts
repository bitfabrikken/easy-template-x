import { ImagePlugin } from './image';
import { LinkPlugin } from './link';
import { LoopPlugin } from './loop';
import { RawXmlPlugin } from './rawXml';
import { TemplatePlugin } from './templatePlugin';
import { TextPlugin } from './text';
import { XlsPlugin } from './xls';

export function createDefaultPlugins(): TemplatePlugin[] {
    return [
        new LoopPlugin(),
        new RawXmlPlugin(),
        new ImagePlugin(),
        new LinkPlugin(),
        new TextPlugin(),
        new XlsPlugin()
    ];
}
