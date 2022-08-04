import { MimeType } from '../../mimeType';
import { Binary } from '../../utils';
import { PluginContent } from '../pluginContent';

export type XlsFormat = MimeType.Xls;

export interface XlsContent extends PluginContent {
    _type: 'xls';
    source: Binary;
    format: XlsFormat;
    width: number;
    height: number;
    dxaOrig: number;
    dyaOrig: number;
    /**
     * Replace a part of the document with raw xml content.
     * If set to `true` the plugin will replace the parent paragraph (<w:p>) of
     * the tag, otherwise it will replace the parent text node (<w:t>).
     */
    replaceParagraph?: boolean;
}
