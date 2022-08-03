import { ScopeData, Tag, TemplateContext } from '../../compilation';
import { MimeTypeHelper } from '../../mimeType';
import { XmlGeneralNode, XmlNode } from '../../xml';
import { TemplatePlugin } from '../templatePlugin';
import { XlsContent } from './xlsContent';

/**
 * Apparently it is not that important for the ID to be unique...
 * Word displays two images correctly even if they both have the same ID.
 * Further more, Word will assign each a unique ID upon saving (it assigns
 * consecutive integers starting with 1).
 *
 * Note: The same principal applies to image names.
 *
 * Tested in Word v1908
 */
let nextFileId = 1;

export class XlsPlugin extends TemplatePlugin {

    public readonly contentType = 'xls';

    public async simpleTagReplacements(tag: Tag, data: ScopeData, context: TemplateContext): Promise<void> {

        const wordTextNode = this.utilities.docxParser.containingTextNode(tag.xmlTextNode);

        const content = data.getScopeData<XlsContent>();
        if (!content || !content.source) {
            XmlNode.remove(wordTextNode);
            return;
        }

        // add the xls file into the archive
        const mediaFilePath = await context.docx.mediaFiles.add(content.source, content.format);
        const relType = MimeTypeHelper.getOfficeRelType(content.format);
        const relId = await context.currentPart.rels.add(mediaFilePath, relType);
        await context.docx.contentTypes.ensureContentType(content.format);

        // create the xml markup
        const fileId = nextFileId++;
        const xlsXml = this.createMarkup(fileId, relId, content.width, content.height);

        XmlNode.insertAfter(xlsXml, wordTextNode);
        XmlNode.remove(wordTextNode);
    }

    private createMarkup(fileId: number, relId: string, width: number, height: number): XmlNode {

        // http://officeopenxml.com/drwPicInline.php

        //
        // Performance note:
        //
        // I've tried to improve the markup generation performance by parsing
        // the string once and caching the result (and of course customizing it
        // per image) but it made no change whatsoever (in both cases 1000 items
        // loop takes around 8 seconds on my machine) so I'm sticking with this
        // approach which I find to be more readable.
        //

        const name = `Xls ${fileId}`;
        const markupText = `
        <w:p>
            <w:pPr>
                <w:pStyle w:val="TextBody"/>
                <w:bidi w:val="0"/>
                <w:jc w:val="left"/>
                <w:rPr>
                    <w:sz w:val="28"/>
                    <w:szCs w:val="28"/>
                </w:rPr>
            </w:pPr>
            <w:r>
                <w:rPr>
                    <w:sz w:val="28"/>
                    <w:szCs w:val="28"/>
                </w:rPr>
                <w:object w:dxaOrig="8974" w:dyaOrig="1280">
                    <v:shape id="ole_${relId}" style="position:absolute;margin-left:2.05pt;margin-top:2.55pt;width:471.3pt;height:58.7pt;mso-position-horizontal-relative:text;mso-position-vertical-relative:text" o:ole="">
                        <v:imagedata r:id="rId3" o:title=""/>
                        <w10:wrap type="square"/>
                    </v:shape>
                    <o:OLEObject Type="Embed" ProgID="Excel.Sheet.12" ShapeID="ole_${relId}" DrawAspect="Content" ObjectID="_319377850" r:id="${relId}"/>
                </w:object>
            </w:r>
        </w:p>
        `;

        const markupXml = this.utilities.xmlParser.parse(markupText) as XmlGeneralNode;
        XmlNode.removeEmptyTextNodes(markupXml); // remove whitespace

        return markupXml;
    }

}
