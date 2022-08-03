import { ScopeData, Tag, TemplateContext } from '../../compilation';
import { MimeTypeHelper } from '../../mimeType';
import { XmlNode } from '../../xml';
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

        const wordTextNode = this.utilities.docxParser.containingParagraphNode(tag.xmlTextNode);

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
        const shapeId = "xls_shape"+relId;
        const progId = "Excel.Sheet.12";

        const wrapType = "topAndBottom"; //or square
        console.log("relId: "+relId+", fileId: "+fileId);

        var xml = `
            <w:p>
                <w:r>
                    <w:object w:dxaOrig="9000" w:dyaOrig="5800">
                        <v:shape id="${shapeId}" style="
                            position:absolute;
                            margin-left:0pt;
                            margin-top:0pt;
                            width:400pt;
                            height:250pt;
                            mso-position-horizontal-relative:text;
                            mso-position-vertical-relative:text;
                        " o:ole="">
                            <w10:wrap type="${wrapType}"/>
                        </v:shape>
                        <o:OLEObject Type="Embed" ProgID="${progId}" ShapeID="${shapeId}" DrawAspect="Content" r:id="${relId}"/>
                    </w:object>
                </w:r>
            </w:p>
        `;
        const newNode = this.utilities.xmlParser.parse(xml);

        XmlNode.insertBefore(newNode, wordTextNode);
        XmlNode.remove(wordTextNode);
    }



}
