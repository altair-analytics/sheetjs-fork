/* 9.3 Relationships */
var RELS = ({
	WB: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
	SHEET: "http://sheetjs.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
	HLINK: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
	VML: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing",
	XPATH: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath",
	XMISS: "http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlPathMissing",
	XLINK: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink",
	CXML: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml",
	CXMLP: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps",
	CMNT: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
	CORE_PROPS: "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
	EXT_PROPS: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties',
	CUST_PROPS: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties',
	SST: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
	STY: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
	THEME: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
	CHART: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
	CHARTEX: "http://schemas.microsoft.com/office/2014/relationships/chartEx",
	CS: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet",
	WS: [
		"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
		"http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet"
	],
	DS: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet",
	MS: "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet",
	IMG: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
	DRAW: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
	XLMETA: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata",
	TCMNT: "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment",
	PEOPLE: "http://schemas.microsoft.com/office/2017/10/relationships/person",
	CONN: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections",
	VBA: "http://schemas.microsoft.com/office/2006/relationships/vbaProject"
}/*:any*/);

/* 9.3.3 Representing Relationships */
function get_rels_path(file/*:string*/)/*:string*/ {
	var n = file.lastIndexOf("/");
	return file.slice(0,n+1) + '_rels/' + file.slice(n+1) + ".rels";
}

function parse_rels(data/*:?string*/, currentFilePath/*:string*/) {
	var rels = {"!id":{}};
	if (!data) return rels;
	if (currentFilePath.charAt(0) !== '/') {
		currentFilePath = '/'+currentFilePath;
	}
	var hash = {};

	(data.match(tagregex)||[]).forEach(function(x) {
		var y = parsexmltag(x);
		/* 9.3.2.2 OPC_Relationships */
		if (y[0] === '<Relationship') {
			var rel = {}; rel.Type = y.Type; rel.Target = unescapexml(y.Target); rel.Id = y.Id; if(y.TargetMode) rel.TargetMode = y.TargetMode;
			var canonictarget = y.TargetMode === 'External' ? y.Target : resolve_path(y.Target, currentFilePath);
			rels[canonictarget] = rel;
			hash[y.Id] = rel;
		}
	});
	rels["!id"] = hash;
	return rels;
}


/* TODO */
function write_rels(rels)/*:string*/ {
	var o = [XML_HEADER, writextag('Relationships', null, {
		//'xmlns:ns0': XMLNS.RELS,
		'xmlns': XMLNS.RELS
	})];
	keys(rels['!id']).forEach(function(rid) {
		o[o.length] = (writextag('Relationship', null, rels['!id'][rid]));
	});
	if(o.length>2){ o[o.length] = ('</Relationships>'); o[1]=o[1].replace("/>",">"); }
	return o.join("");
}

var DRAW_ROOT = writextag('xdr:wsDr', null, {
	'xmlns:xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
	'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
	//'xmlns:ns0': XMLNS.RELS,
	// 'xmlns': XMLNS.RELS
});

function write_drawing(images, worksheet) {
	var o = [];
	o.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
	o.push('<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">');

	images.forEach((image, i) => {
		const pos = image.position || {};
		const id = i + 1;
		const name = image.name || `Image${id}`;
		const attrs = image.attrs || { editAs: "oneCell" };

		let anchor = '';
		if (pos.type === 'twoCellAnchor') {
			const from = pos.from || { col: 0, row: 0, colOff: 0, rowOff: 0 };
			const to = pos.to || { col: 0, row: 0, colOff: 0, rowOff: 0 };

			anchor = `
        <xdr:from>
          <xdr:col>${from.col}</xdr:col>
          <xdr:colOff>${from.colOff}</xdr:colOff>
          <xdr:row>${from.row}</xdr:row>
          <xdr:rowOff>${from.rowOff}</xdr:rowOff>
        </xdr:from>
        <xdr:to>
          <xdr:col>${to.col}</xdr:col>
          <xdr:colOff>${to.colOff}</xdr:colOff>
          <xdr:row>${to.row}</xdr:row>
          <xdr:rowOff>${to.rowOff}</xdr:rowOff>
        </xdr:to>
      `;
		} else if (pos.type === 'centeredInCell') {
			// New positioning type for centered images
			const centerPos = drawingCalculateCenteredPosition(pos, worksheet, image);

			anchor = `
        <xdr:from>
          <xdr:col>${centerPos.from.col}</xdr:col>
          <xdr:colOff>${centerPos.from.colOff}</xdr:colOff>
          <xdr:row>${centerPos.from.row}</xdr:row>
          <xdr:rowOff>${centerPos.from.rowOff}</xdr:rowOff>
        </xdr:from>
        <xdr:to>
          <xdr:col>${centerPos.to.col}</xdr:col>
          <xdr:colOff>${centerPos.to.colOff}</xdr:colOff>
          <xdr:row>${centerPos.to.row}</xdr:row>
          <xdr:rowOff>${centerPos.to.rowOff}</xdr:rowOff>
        </xdr:to>
      `;
		}

		const pic = `
      <xdr:pic>
        <xdr:nvPicPr>
          <xdr:cNvPr id="${id}" name="${name}"/>
          <xdr:cNvPicPr>
            <a:picLocks noChangeAspect="1"/>
          </xdr:cNvPicPr>
        </xdr:nvPicPr>
        <xdr:blipFill>
          <a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId${id}"/>
          <a:stretch>
            <a:fillRect/>
          </a:stretch>
        </xdr:blipFill>
        <xdr:spPr>
          <a:prstGeom prst="rect">
            <a:avLst/>
          </a:prstGeom>
        </xdr:spPr>
      </xdr:pic>
      <xdr:clientData/>
    `;

		const anchorType = pos.type === 'centeredInCell' ? 'twoCellAnchor' : pos.type;
		o.push(`<xdr:${anchorType} editAs="${attrs.editAs || 'oneCell'}">${anchor}${pic}</xdr:${anchorType}>`);
	});

	o.push('</xdr:wsDr>');
	return o.join('\n');
}

function drawingCalculateCenteredPosition(pos, worksheet, image) {
	// Default cell dimensions in EMUs (English Metric Units)
	// 1 pixel ≈ 9525 EMUs, but Excel uses different ratios for rows/cols
	const DEFAULT_COL_WIDTH_EMU = 64 * 9525;
	const DEFAULT_ROW_HEIGHT_EMU = 20 * 9525;

	const colWidth = drawingGetCellWidth(worksheet, pos.col) || DEFAULT_COL_WIDTH_EMU;
	const rowHeight = drawingGetCellHeight(worksheet, pos.row) || DEFAULT_ROW_HEIGHT_EMU;

	const imageWidth = (image.width || 100) * 9525;
	const imageHeight = (image.height || 80) * 9525;

	const padding = (pos.padding || 5) * 9525;

	const availableWidth = colWidth - (2 * padding);
	const availableHeight = rowHeight - (2 * padding);

	const horizontalOffset = padding + (availableWidth - imageWidth) / 2;
	const verticalOffset = padding + (availableHeight - imageHeight) / 2;

	const finalImageWidth = Math.min(imageWidth, availableWidth);
	const finalImageHeight = Math.min(imageHeight, availableHeight);

	return {
		from: {
			col: pos.col,
			row: pos.row,
			colOff: Math.max(0, horizontalOffset),
			rowOff: Math.max(0, verticalOffset)
		},
		to: {
			col: pos.col,
			row: pos.row,
			colOff: Math.max(0, horizontalOffset + finalImageWidth),
			rowOff: Math.max(0, verticalOffset + finalImageHeight)
		}
	};
}

function drawingGetCellWidth(worksheet, col) {
	if (worksheet['!cols'] && worksheet['!cols'][col] && worksheet['!cols'][col].width) {
		return worksheet['!cols'][col].width * 7 * 9525;
	}
	return null;
}

function drawingGetCellHeight(worksheet, row) {
	if (worksheet['!rows'] && worksheet['!rows'][row] && worksheet['!rows'][row].hpt) {
		return worksheet['!rows'][row].hpt * 12700;
	}
	return null;
}

function add_rels(rels, rId/*:number*/, f, type, relobj, targetmode/*:?string*/)/*:number*/ {
	if(!relobj) relobj = {};
	if(!rels['!id']) rels['!id'] = {};
	if(!rels['!idx']) rels['!idx'] = 1;
	if(rId < 0) for(rId = rels['!idx']; rels['!id']['rId' + rId]; ++rId){/* empty */}
	rels['!idx'] = rId + 1;
	relobj.Id = 'rId' + rId;
	relobj.Type = type;
	relobj.Target = f;
	if(targetmode) relobj.TargetMode = targetmode;
	else if([RELS.HLINK, RELS.XPATH, RELS.XMISS].indexOf(relobj.Type) > -1) relobj.TargetMode = "External";
	if(rels['!id'][relobj.Id]) throw new Error("Cannot rewrite rId " + rId);
	rels['!id'][relobj.Id] = relobj;
	rels[('/' + relobj.Target).replace("//","/")] = relobj;
	return rId;
}
