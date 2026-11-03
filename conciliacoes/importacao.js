/**
 * Conciliações Bancárias — NOCTUS. Parte 2 (02/11/2026): planilha de importação
 * dos PIX recebidos no sistema contábil.
 *
 * Regras do Rocha (02/11/2026):
 *   - SÓ os PIX recebidos do extrato OFX (entrada, histórico "PIX RECEBIDO"),
 *     na data da conciliação. Nada de PIX enviado, pagamento, tarifa ou linha
 *     de saldo (SALDO ANTERIOR / SALDO TOTAL DISPONÍVEL DIA).
 *   - Débito (conta) = o código da conta no SCK do relatório importado (1001283
 *     na Praia Verde); Valor e Data exatamente como no OFX; Histórico no formato
 *     "PIX RECEBIDO 30/10 - AFE GESTAO COMERCIAL 40.248.519/0001-08".
 *   - Layout EXATO da planilha modelo "Modelo planilha de importação - Depósitos":
 *     aba "Planilha1", colunas Débito | Valor | Data | Histórico, mesmos
 *     estilos, larguras, filtro e sem linhas de grade.
 *
 * Por que o arquivo é montado à mão (e não pelo SheetJS): a versão gratuita do
 * SheetJS não grava estilos, e o modelo tem formato contábil no valor, data
 * curta, bordas e cabeçalho colorido. Aqui as peças do próprio modelo
 * (estilos, tema, workbook) vão embutidas e só as linhas são geradas, então o
 * arquivo sai igual ao modelo. O .xlsx é um ZIP; as peças vão sem compressão
 * ("stored"), que o Excel aceita normalmente.
 *
 * Peças geradas a partir do modelo por gerador/gerar_importacao.py, com o
 * caminho local do autor e os nomes do docProps removidos (repositório público).
 */
import { historicoImportacao, tipoExtrato } from "./composicao.js";

const PECAS = {
"[Content_Types].xml": "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/><Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/><Override PartName=\"/xl/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/><Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/><Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/><Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/><Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/></Types>",
"_rels/.rels": "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>",
"docProps/app.xml": "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\"><Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size=\"2\" baseType=\"variant\"><vt:variant><vt:lpstr>Folhas de Cálculo</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size=\"1\" baseType=\"lpstr\"><vt:lpstr>Planilha1</vt:lpstr></vt:vector></TitlesOfParts><Company></Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0300</AppVersion></Properties>",
"docProps/core.xml": "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"><dc:creator>NOCTUS</dc:creator><cp:lastModifiedBy>NOCTUS</cp:lastModifiedBy><dcterms:created xsi:type=\"dcterms:W3CDTF\">@@CRIADO@@</dcterms:created><dcterms:modified xsi:type=\"dcterms:W3CDTF\">@@CRIADO@@</dcterms:modified></cp:coreProperties>",
"xl/workbook.xml": "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x15 xr xr6 xr10 xr2\" xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main\" xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\" xmlns:xr6=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision6\" xmlns:xr10=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision10\" xmlns:xr2=\"http://schemas.microsoft.com/office/spreadsheetml/2015/revision2\"><fileVersion appName=\"xl\" lastEdited=\"7\" lowestEdited=\"7\" rupBuild=\"27932\"/><workbookPr defaultThemeVersion=\"166925\"/><xr:revisionPtr revIDLastSave=\"0\" documentId=\"13_ncr:1_{668475A6-28D5-4BE9-BE8F-55152B5F0553}\" xr6:coauthVersionLast=\"47\" xr6:coauthVersionMax=\"47\" xr10:uidLastSave=\"{00000000-0000-0000-0000-000000000000}\"/><bookViews><workbookView xWindow=\"-98\" yWindow=\"-98\" windowWidth=\"21795\" windowHeight=\"12975\" xr2:uid=\"{00000000-000D-0000-FFFF-FFFF00000000}\"/></bookViews><sheets><sheet name=\"Planilha1\" sheetId=\"1\" r:id=\"rId1\"/></sheets><definedNames><definedName name=\"_xlnm._FilterDatabase\" localSheetId=\"0\" hidden=\"1\">Planilha1!$A$1:$D$2</definedName></definedNames><calcPr calcId=\"191029\"/><extLst><ext uri=\"{140A7094-0E35-4892-8432-C4D2E57EDEB5}\" xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main\"><x15:workbookPr chartTrackingRefBase=\"1\"/></ext><ext uri=\"{B58B0392-4F1F-4190-BB64-5DF3571DCE5F}\" xmlns:xcalcf=\"http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures\"><xcalcf:calcFeatures><xcalcf:feature name=\"microsoft.com:RD\"/><xcalcf:feature name=\"microsoft.com:Single\"/><xcalcf:feature name=\"microsoft.com:FV\"/><xcalcf:feature name=\"microsoft.com:CNMTM\"/><xcalcf:feature name=\"microsoft.com:LET_WF\"/><xcalcf:feature name=\"microsoft.com:LAMBDA_WF\"/><xcalcf:feature name=\"microsoft.com:ARRAYTEXT_WF\"/></xcalcf:calcFeatures></ext><ext uri=\"{D14903EA-33C4-47F7-8F05-3474C54BE107}\" xmlns:xlwcv=\"http://schemas.microsoft.com/office/spreadsheetml/2024/workbookCompatibilityVersion\"><xlwcv:version setVersion=\"1\"/></ext></extLst></workbook>",
"xl/_rels/workbook.xml.rels": "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/><Relationship Id=\"rId4\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/></Relationships>",
"xl/styles.xml": "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac x16r2 xr\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\" xmlns:x16r2=\"http://schemas.microsoft.com/office/spreadsheetml/2015/02/main\" xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\"><numFmts count=\"1\"><numFmt numFmtId=\"44\" formatCode=\"_-&quot;R$&quot;\\ * #,##0.00_-;\\-&quot;R$&quot;\\ * #,##0.00_-;_-&quot;R$&quot;\\ * &quot;-&quot;??_-;_-@_-\"/></numFmts><fonts count=\"6\" x14ac:knownFonts=\"1\"><font><sz val=\"11\"/><color theme=\"1\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font><font><sz val=\"11\"/><color rgb=\"FF000000\"/><name val=\"Calibri\"/><family val=\"2\"/></font><font><sz val=\"8\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font><font><sz val=\"10\"/><name val=\"Arial\"/><family val=\"2\"/></font><font><sz val=\"10\"/><name val=\"Arial\"/><family val=\"2\"/></font><font><sz val=\"11\"/><color indexed=\"8\"/><name val=\"Calibri\"/><family val=\"2\"/></font></fonts><fills count=\"3\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill><fill><patternFill patternType=\"solid\"><fgColor theme=\"4\" tint=\"0.79998168889431442\"/><bgColor indexed=\"64\"/></patternFill></fill></fills><borders count=\"3\"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style=\"thin\"><color indexed=\"64\"/></left><right style=\"thin\"><color indexed=\"64\"/></right><top style=\"thin\"><color indexed=\"64\"/></top><bottom style=\"thin\"><color indexed=\"64\"/></bottom><diagonal/></border><border><left style=\"thin\"><color indexed=\"64\"/></left><right/><top style=\"thin\"><color indexed=\"64\"/></top><bottom style=\"thin\"><color indexed=\"64\"/></bottom><diagonal/></border></borders><cellStyleXfs count=\"3\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/><xf numFmtId=\"0\" fontId=\"3\" fillId=\"0\" borderId=\"0\"/><xf numFmtId=\"0\" fontId=\"4\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs><cellXfs count=\"11\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/><xf numFmtId=\"0\" fontId=\"0\" fillId=\"2\" borderId=\"1\" xfId=\"0\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"center\"/></xf><xf numFmtId=\"44\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\" applyAlignment=\"1\"><alignment horizontal=\"right\"/></xf><xf numFmtId=\"44\" fontId=\"0\" fillId=\"2\" borderId=\"1\" xfId=\"0\" applyNumberFormat=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"center\"/></xf><xf numFmtId=\"0\" fontId=\"0\" fillId=\"2\" borderId=\"1\" xfId=\"0\" applyFill=\"1\" applyBorder=\"1\"/><xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"1\" xfId=\"0\" applyFont=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"center\" vertical=\"center\"/></xf><xf numFmtId=\"0\" fontId=\"0\" fillId=\"2\" borderId=\"1\" xfId=\"0\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"center\" vertical=\"center\"/></xf><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyAlignment=\"1\"><alignment horizontal=\"center\" vertical=\"center\"/></xf><xf numFmtId=\"44\" fontId=\"5\" fillId=\"0\" borderId=\"1\" xfId=\"0\" applyNumberFormat=\"1\" applyFont=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"center\" vertical=\"center\"/></xf><xf numFmtId=\"14\" fontId=\"5\" fillId=\"0\" borderId=\"2\" xfId=\"0\" applyNumberFormat=\"1\" applyFont=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"center\" vertical=\"center\"/></xf><xf numFmtId=\"0\" fontId=\"5\" fillId=\"0\" borderId=\"1\" xfId=\"0\" applyFont=\"1\" applyBorder=\"1\"/></cellXfs><cellStyles count=\"3\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/><cellStyle name=\"Normal 2\" xfId=\"1\" xr:uid=\"{3A47B5E3-307C-435D-A9DA-48FB7C5673CC}\"/><cellStyle name=\"Normal 3\" xfId=\"2\" xr:uid=\"{D63F7722-49C1-413C-9977-68E8475CF5FC}\"/></cellStyles><dxfs count=\"0\"/><tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium2\" defaultPivotStyle=\"PivotStyleLight16\"/><extLst><ext uri=\"{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}\" xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"><x14:slicerStyles defaultSlicerStyle=\"SlicerStyleLight1\"/></ext><ext uri=\"{9260A510-F301-46a8-8635-F512D64BE5F5}\" xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main\"><x15:timelineStyles defaultTimelineStyle=\"TimeSlicerStyleLight1\"/></ext></extLst></styleSheet>",
"xl/theme/theme1.xml": "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Tema Office 2013 - 2022\"><a:themeElements><a:clrScheme name=\"Office 2013 - 2022\"><a:dk1><a:sysClr val=\"windowText\" lastClr=\"000000\"/></a:dk1><a:lt1><a:sysClr val=\"window\" lastClr=\"FFFFFF\"/></a:lt1><a:dk2><a:srgbClr val=\"44546A\"/></a:dk2><a:lt2><a:srgbClr val=\"E7E6E6\"/></a:lt2><a:accent1><a:srgbClr val=\"4472C4\"/></a:accent1><a:accent2><a:srgbClr val=\"ED7D31\"/></a:accent2><a:accent3><a:srgbClr val=\"A5A5A5\"/></a:accent3><a:accent4><a:srgbClr val=\"FFC000\"/></a:accent4><a:accent5><a:srgbClr val=\"5B9BD5\"/></a:accent5><a:accent6><a:srgbClr val=\"70AD47\"/></a:accent6><a:hlink><a:srgbClr val=\"0563C1\"/></a:hlink><a:folHlink><a:srgbClr val=\"954F72\"/></a:folHlink></a:clrScheme><a:fontScheme name=\"Office 2013 - 2022\"><a:majorFont><a:latin typeface=\"Calibri Light\" panose=\"020F0302020204030204\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/><a:font script=\"Jpan\" typeface=\"游ゴシック Light\"/><a:font script=\"Hang\" typeface=\"맑은 고딕\"/><a:font script=\"Hans\" typeface=\"等线 Light\"/><a:font script=\"Hant\" typeface=\"新細明體\"/><a:font script=\"Arab\" typeface=\"Times New Roman\"/><a:font script=\"Hebr\" typeface=\"Times New Roman\"/><a:font script=\"Thai\" typeface=\"Tahoma\"/><a:font script=\"Ethi\" typeface=\"Nyala\"/><a:font script=\"Beng\" typeface=\"Vrinda\"/><a:font script=\"Gujr\" typeface=\"Shruti\"/><a:font script=\"Khmr\" typeface=\"MoolBoran\"/><a:font script=\"Knda\" typeface=\"Tunga\"/><a:font script=\"Guru\" typeface=\"Raavi\"/><a:font script=\"Cans\" typeface=\"Euphemia\"/><a:font script=\"Cher\" typeface=\"Plantagenet Cherokee\"/><a:font script=\"Yiii\" typeface=\"Microsoft Yi Baiti\"/><a:font script=\"Tibt\" typeface=\"Microsoft Himalaya\"/><a:font script=\"Thaa\" typeface=\"MV Boli\"/><a:font script=\"Deva\" typeface=\"Mangal\"/><a:font script=\"Telu\" typeface=\"Gautami\"/><a:font script=\"Taml\" typeface=\"Latha\"/><a:font script=\"Syrc\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Orya\" typeface=\"Kalinga\"/><a:font script=\"Mlym\" typeface=\"Kartika\"/><a:font script=\"Laoo\" typeface=\"DokChampa\"/><a:font script=\"Sinh\" typeface=\"Iskoola Pota\"/><a:font script=\"Mong\" typeface=\"Mongolian Baiti\"/><a:font script=\"Viet\" typeface=\"Times New Roman\"/><a:font script=\"Uigh\" typeface=\"Microsoft Uighur\"/><a:font script=\"Geor\" typeface=\"Sylfaen\"/><a:font script=\"Armn\" typeface=\"Arial\"/><a:font script=\"Bugi\" typeface=\"Leelawadee UI\"/><a:font script=\"Bopo\" typeface=\"Microsoft JhengHei\"/><a:font script=\"Java\" typeface=\"Javanese Text\"/><a:font script=\"Lisu\" typeface=\"Segoe UI\"/><a:font script=\"Mymr\" typeface=\"Myanmar Text\"/><a:font script=\"Nkoo\" typeface=\"Ebrima\"/><a:font script=\"Olck\" typeface=\"Nirmala UI\"/><a:font script=\"Osma\" typeface=\"Ebrima\"/><a:font script=\"Phag\" typeface=\"Phagspa\"/><a:font script=\"Syrn\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Syrj\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Syre\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Sora\" typeface=\"Nirmala UI\"/><a:font script=\"Tale\" typeface=\"Microsoft Tai Le\"/><a:font script=\"Talu\" typeface=\"Microsoft New Tai Lue\"/><a:font script=\"Tfng\" typeface=\"Ebrima\"/></a:majorFont><a:minorFont><a:latin typeface=\"Calibri\" panose=\"020F0502020204030204\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/><a:font script=\"Jpan\" typeface=\"游ゴシック\"/><a:font script=\"Hang\" typeface=\"맑은 고딕\"/><a:font script=\"Hans\" typeface=\"等线\"/><a:font script=\"Hant\" typeface=\"新細明體\"/><a:font script=\"Arab\" typeface=\"Arial\"/><a:font script=\"Hebr\" typeface=\"Arial\"/><a:font script=\"Thai\" typeface=\"Tahoma\"/><a:font script=\"Ethi\" typeface=\"Nyala\"/><a:font script=\"Beng\" typeface=\"Vrinda\"/><a:font script=\"Gujr\" typeface=\"Shruti\"/><a:font script=\"Khmr\" typeface=\"DaunPenh\"/><a:font script=\"Knda\" typeface=\"Tunga\"/><a:font script=\"Guru\" typeface=\"Raavi\"/><a:font script=\"Cans\" typeface=\"Euphemia\"/><a:font script=\"Cher\" typeface=\"Plantagenet Cherokee\"/><a:font script=\"Yiii\" typeface=\"Microsoft Yi Baiti\"/><a:font script=\"Tibt\" typeface=\"Microsoft Himalaya\"/><a:font script=\"Thaa\" typeface=\"MV Boli\"/><a:font script=\"Deva\" typeface=\"Mangal\"/><a:font script=\"Telu\" typeface=\"Gautami\"/><a:font script=\"Taml\" typeface=\"Latha\"/><a:font script=\"Syrc\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Orya\" typeface=\"Kalinga\"/><a:font script=\"Mlym\" typeface=\"Kartika\"/><a:font script=\"Laoo\" typeface=\"DokChampa\"/><a:font script=\"Sinh\" typeface=\"Iskoola Pota\"/><a:font script=\"Mong\" typeface=\"Mongolian Baiti\"/><a:font script=\"Viet\" typeface=\"Arial\"/><a:font script=\"Uigh\" typeface=\"Microsoft Uighur\"/><a:font script=\"Geor\" typeface=\"Sylfaen\"/><a:font script=\"Armn\" typeface=\"Arial\"/><a:font script=\"Bugi\" typeface=\"Leelawadee UI\"/><a:font script=\"Bopo\" typeface=\"Microsoft JhengHei\"/><a:font script=\"Java\" typeface=\"Javanese Text\"/><a:font script=\"Lisu\" typeface=\"Segoe UI\"/><a:font script=\"Mymr\" typeface=\"Myanmar Text\"/><a:font script=\"Nkoo\" typeface=\"Ebrima\"/><a:font script=\"Olck\" typeface=\"Nirmala UI\"/><a:font script=\"Osma\" typeface=\"Ebrima\"/><a:font script=\"Phag\" typeface=\"Phagspa\"/><a:font script=\"Syrn\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Syrj\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Syre\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Sora\" typeface=\"Nirmala UI\"/><a:font script=\"Tale\" typeface=\"Microsoft Tai Le\"/><a:font script=\"Talu\" typeface=\"Microsoft New Tai Lue\"/><a:font script=\"Tfng\" typeface=\"Ebrima\"/></a:minorFont></a:fontScheme><a:fmtScheme name=\"Office 2013 - 2022\"><a:fillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"110000\"/><a:satMod val=\"105000\"/><a:tint val=\"67000\"/></a:schemeClr></a:gs><a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"105000\"/><a:satMod val=\"103000\"/><a:tint val=\"73000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"105000\"/><a:satMod val=\"109000\"/><a:tint val=\"81000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:satMod val=\"103000\"/><a:lumMod val=\"102000\"/><a:tint val=\"94000\"/></a:schemeClr></a:gs><a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:satMod val=\"110000\"/><a:lumMod val=\"100000\"/><a:shade val=\"100000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"99000\"/><a:satMod val=\"120000\"/><a:shade val=\"78000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w=\"6350\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln><a:ln w=\"12700\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln><a:ln w=\"19050\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad=\"57150\" dist=\"19050\" dir=\"5400000\" algn=\"ctr\" rotWithShape=\"0\"><a:srgbClr val=\"000000\"><a:alpha val=\"63000\"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:solidFill><a:schemeClr val=\"phClr\"><a:tint val=\"95000\"/><a:satMod val=\"170000\"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:tint val=\"93000\"/><a:satMod val=\"150000\"/><a:shade val=\"98000\"/><a:lumMod val=\"102000\"/></a:schemeClr></a:gs><a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:tint val=\"98000\"/><a:satMod val=\"130000\"/><a:shade val=\"90000\"/><a:lumMod val=\"103000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:shade val=\"63000\"/><a:satMod val=\"120000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/><a:extLst><a:ext uri=\"{05A4C25C-085E-4340-85A3-A5531E510DB2}\"><thm15:themeFamily xmlns:thm15=\"http://schemas.microsoft.com/office/thememl/2012/main\" name=\"Office 2013 - 2022 Theme\" id=\"{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}\" vid=\"{4A3C46E8-61CC-4603-A589-7422A47A8E4A}\"/></a:ext></a:extLst></a:theme>"
};
const CAB_PLANILHA = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac xr xr2 xr3\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\" xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\" xmlns:xr2=\"http://schemas.microsoft.com/office/spreadsheetml/2015/revision2\" xmlns:xr3=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision3\" xr:uid=\"{00000000-0001-0000-0000-000000000000}\"><dimension ref=\"@@REF@@\"/><sheetViews><sheetView showGridLines=\"0\" tabSelected=\"1\" zoomScaleNormal=\"100\" workbookViewId=\"0\"><selection activeCell=\"A2\" sqref=\"A2\"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight=\"14.25\" x14ac:dyDescent=\"0.45\"/><cols><col min=\"1\" max=\"1\" width=\"8\" bestFit=\"1\" customWidth=\"1\"/><col min=\"2\" max=\"2\" width=\"15.53125\" style=\"2\" bestFit=\"1\" customWidth=\"1\"/><col min=\"3\" max=\"3\" width=\"12.86328125\" style=\"7\" bestFit=\"1\" customWidth=\"1\"/><col min=\"4\" max=\"4\" width=\"126.33203125\" bestFit=\"1\" customWidth=\"1\"/><col min=\"8\" max=\"8\" width=\"13.86328125\" customWidth=\"1\"/><col min=\"9\" max=\"9\" width=\"27.53125\" customWidth=\"1\"/><col min=\"10\" max=\"10\" width=\"19.46484375\" bestFit=\"1\" customWidth=\"1\"/></cols>";
const RODAPE_PLANILHA = "<autoFilter ref=\"@@REF@@\" xr:uid=\"{00000000-0001-0000-0000-000000000000}\"/><phoneticPr fontId=\"2\" type=\"noConversion\"/><pageMargins left=\"0.511811024\" right=\"0.511811024\" top=\"0.78740157499999996\" bottom=\"0.78740157499999996\" header=\"0.31496062000000002\" footer=\"0.31496062000000002\"/><pageSetup paperSize=\"9\" orientation=\"portrait\"/></worksheet>";
const LINHA_CABECALHO = "<row r=\"1\" spans=\"1:4\" x14ac:dyDescent=\"0.45\"><c r=\"A1\" s=\"1\" t=\"s\"><v>0</v></c><c r=\"B1\" s=\"3\" t=\"s\"><v>1</v></c><c r=\"C1\" s=\"6\" t=\"s\"><v>2</v></c><c r=\"D1\" s=\"4\" t=\"s\"><v>3</v></c></row>";
const ATRIBUTOS_LINHA = " spans=\"1:4\" x14ac:dyDescent=\"0.45\"";
const ESTILO = {"A": "5", "B": "8", "C": "9", "D": "10"};
const CABECALHOS = ["Débito", "Valor", "Data", "Histórico"];

const r2 = (n) => Math.round(Number(n || 0) * 100) / 100;
function aviso(msg) { const e = new Error(msg); e.aviso = true; return e; }

/**
 * Linhas da planilha a partir do relatório do SCK (conta e data) e do extrato.
 * Devolve { linhas:[{conta, valor, data, historico, historicoOriginal, padrao, dataConfere, fitid}],
 *           total, alertas:[] }.
 */
export function linhasImportacao({ sck = null, ofx = null } = {}) {
  if (!sck || !ofx) throw aviso("Importe o relatório do SCK e o extrato antes de gerar os lançamentos.");
  const conta = Number(String(sck.conta?.codigoSck || "").replace(/\D/g, ""));
  if (!conta) throw aviso("O relatório do SCK não informa o código da conta; sem ele não dá para gerar a planilha.");
  const dataAlvo = sck.data?.iso || null;

  const linhas = (ofx.lancamentos || [])
    .filter((t) => t.direcao === "entrada" && tipoExtrato(t.memo) === "pixRecebido")
    .filter((t) => t.data && (!dataAlvo || t.data.iso === dataAlvo))
    .map((t) => {
      const h = historicoImportacao(t.memo, t.data);
      return { conta, valor: r2(t.valor), data: t.data, historico: h.texto, historicoOriginal: t.memo,
               padrao: h.padrao, dataConfere: h.dataConfere, fitid: t.fitid || "" };
    });

  const alertas = [];
  const fora = linhas.filter((l) => !l.padrao);
  if (fora.length) {
    alertas.push(`${fora.length} PIX recebido(s) com histórico fora do padrão: foram mantidos como vieram no extrato. Confira antes de importar.`);
  }
  const dataDiferente = linhas.filter((l) => l.dataConfere === false);
  if (dataDiferente.length) {
    alertas.push(`${dataDiferente.length} PIX recebido(s) têm no histórico um dia diferente da data do lançamento no extrato. Confira antes de importar.`);
  }
  return { linhas, total: r2(linhas.reduce((s, l) => s + l.valor, 0)), alertas };
}

/** Número de série do Excel (dia 0 = 30/12/1899) para a data ISO, sem fuso. */
export function serialExcel(iso) {
  const [a, m, d] = String(iso).split("-").map(Number);
  return Math.round(Date.UTC(a, m - 1, d) / 86400000) + 25569;
}

const escXml = (t) => String(t ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]))
  // caractere de controle não é válido em XML 1.0 (um histórico estranho não pode corromper o arquivo)
  .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "");

/** Monta as peças XML da planilha (exportado para os testes conferirem o conteúdo). */
export function pecasPlanilha(linhas, agora = new Date()) {
  const textos = [...CABECALHOS];
  const indice = new Map(textos.map((t, i) => [t, i]));
  const idTexto = (t) => {
    if (!indice.has(t)) { indice.set(t, textos.length); textos.push(t); }
    return indice.get(t);
  };
  let usos = CABECALHOS.length;
  const ultima = linhas.length + 1;
  const ref = `A1:D${Math.max(ultima, 2)}`;

  const corpo = linhas.map((l, i) => {
    const r = i + 2;
    usos++;
    return `<row r="${r}"${ATRIBUTOS_LINHA}>` +
      `<c r="A${r}" s="${ESTILO.A}"><v>${l.conta}</v></c>` +
      `<c r="B${r}" s="${ESTILO.B}"><v>${r2(l.valor)}</v></c>` +
      `<c r="C${r}" s="${ESTILO.C}"><v>${serialExcel(l.data.iso)}</v></c>` +
      `<c r="D${r}" s="${ESTILO.D}" t="s"><v>${idTexto(l.historico)}</v></c></row>`;
  }).join("");

  const sheet = CAB_PLANILHA.replace("@@REF@@", ref) + "<sheetData>" + LINHA_CABECALHO + corpo + "</sheetData>" +
    RODAPE_PLANILHA.replace("@@REF@@", ref);
  const sst = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
    `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${usos}" uniqueCount="${textos.length}">` +
    textos.map((t) => `<si><t>${escXml(t)}</t></si>`).join("") + "</sst>";
  const workbook = PECAS["xl/workbook.xml"].replace(/Planilha1!\$A\$1:\$D\$\d+/, `Planilha1!$A$1:$D$${Math.max(ultima, 2)}`);
  const criado = agora.toISOString().replace(/\.\d{3}Z$/, "Z");

  return {
    "[Content_Types].xml": PECAS["[Content_Types].xml"],
    "_rels/.rels": PECAS["_rels/.rels"],
    "docProps/app.xml": PECAS["docProps/app.xml"],
    "docProps/core.xml": PECAS["docProps/core.xml"].replaceAll("@@CRIADO@@", criado),
    "xl/workbook.xml": workbook,
    "xl/_rels/workbook.xml.rels": PECAS["xl/_rels/workbook.xml.rels"],
    "xl/styles.xml": PECAS["xl/styles.xml"],
    "xl/theme/theme1.xml": PECAS["xl/theme/theme1.xml"],
    "xl/worksheets/sheet1.xml": sheet,
    "xl/sharedStrings.xml": sst,
  };
}

/* ------------------------------------------------------------ ZIP (sem compressão) */
const TABELA_CRC = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();
function crc32(bytes) {
  let c = 0xFFFFFFFF;
  for (let i = 0; i < bytes.length; i++) c = TABELA_CRC[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
  return (c ^ 0xFFFFFFFF) >>> 0;
}
function zipSemCompressao(arquivos) {
  const enc = new TextEncoder();
  const locais = [], central = [];
  let deslocamento = 0;
  for (const [nome, conteudo] of Object.entries(arquivos)) {
    const n = enc.encode(nome), d = typeof conteudo === "string" ? enc.encode(conteudo) : conteudo;
    const crc = crc32(d);
    const loc = new DataView(new ArrayBuffer(30));
    loc.setUint32(0, 0x04034b50, true); loc.setUint16(4, 20, true); loc.setUint16(6, 0x0800, true);
    loc.setUint16(8, 0, true); loc.setUint16(10, 0, true); loc.setUint16(12, 0x21, true);
    loc.setUint32(14, crc, true); loc.setUint32(18, d.length, true); loc.setUint32(22, d.length, true);
    loc.setUint16(26, n.length, true); loc.setUint16(28, 0, true);
    locais.push(new Uint8Array(loc.buffer), n, d);
    const cen = new DataView(new ArrayBuffer(46));
    cen.setUint32(0, 0x02014b50, true); cen.setUint16(4, 20, true); cen.setUint16(6, 20, true);
    cen.setUint16(8, 0x0800, true); cen.setUint16(10, 0, true); cen.setUint16(12, 0, true); cen.setUint16(14, 0x21, true);
    cen.setUint32(16, crc, true); cen.setUint32(20, d.length, true); cen.setUint32(24, d.length, true);
    cen.setUint16(28, n.length, true); cen.setUint32(42, deslocamento, true);
    central.push(new Uint8Array(cen.buffer), n);
    deslocamento += 30 + n.length + d.length;
  }
  const tamCentral = central.reduce((s, p) => s + p.length, 0);
  const fim = new DataView(new ArrayBuffer(22));
  fim.setUint32(0, 0x06054b50, true);
  const qtd = Object.keys(arquivos).length;
  fim.setUint16(8, qtd, true); fim.setUint16(10, qtd, true);
  fim.setUint32(12, tamCentral, true); fim.setUint32(16, deslocamento, true);
  const partes = [...locais, ...central, new Uint8Array(fim.buffer)];
  const saida = new Uint8Array(partes.reduce((s, p) => s + p.length, 0));
  let pos = 0;
  for (const p of partes) { saida.set(p, pos); pos += p.length; }
  return saida;
}

/** Bytes do .xlsx pronto para importar. */
export function planilhaImportacao(linhas, agora = new Date()) {
  if (!linhas?.length) throw aviso("Não há PIX recebido no extrato desta data: nada a gerar.");
  return zipSemCompressao(pecasPlanilha(linhas, agora));
}

/** "Importacao_PIX_1001283_2026-10-30.xlsx" */
export const nomeArquivoImportacao = (conta, dataIso) => `Importacao_PIX_${conta}_${dataIso}.xlsx`;
