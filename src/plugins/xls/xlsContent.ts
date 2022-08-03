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
}
