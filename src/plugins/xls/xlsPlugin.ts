import { ScopeData, Tag, TemplateContext } from '../../compilation';
import { MimeTypeHelper } from '../../mimeType';
import { XmlNode } from '../../xml';
import { TemplatePlugin } from '../templatePlugin';
import { XlsContent } from './xlsContent';


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
        // for now, for brevity, we add files to mediaDir (word/media in the xls archive)
        const mediaFilePath = await context.docx.mediaFiles.add(content.source, content.format);
        const relType = MimeTypeHelper.getOfficeRelType(content.format);
        const relId = await context.currentPart.rels.add(mediaFilePath, relType);
        await context.docx.contentTypes.ensureContentType(content.format);

        
        // prepare vars for xml markup
        const shapeId = "xls_shape"+relId;
        const progId = "Excel.Sheet.12"; //not sure why this is needed yet or if it'll cause any trouble leaving like this
        const wrapType = "topAndBottom"; //or square

        const width = content.width || 450;
        const height = content.height || 300;
        const dxaOrig = content.dxaOrig || 9334; //no time atm., to figure out how this parameter works
        const dyaOrig = content.dyaOrig || 5800; //no time atm., to figure out how this parameter works


        // create the xml markup
        var xml = `
            <w:p>
                <w:pPr>
                    <w:pStyle w:val="Normal"/>
                    <w:rPr/>
                </w:pPr>
                <w:r>
                    <w:object w:dxaOrig="${dxaOrig}" w:dyaOrig="${dyaOrig}">
                        <v:shape id="${shapeId}" style="
                            position:absolute;
                            margin-left:0pt;
                            margin-top:0pt;
                            width:${width};
                            height:${height};
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
