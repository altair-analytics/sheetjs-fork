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

function write_drawing(images) {
	var o = [];
	o[o.length] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
	o[o.length] = '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">';

	for (var i = 0; i < images.length; i++) {
		var image = images[i];
		var pos = image.position || {};

		if (pos.type === 'twoCellAnchor') {
			var from = pos.from || {}, to = pos.to || {},
				fromCol = from.col || 0, toCol = to.col || 0,
				fromRow = from.row || 0, toRow = to.row || 0;

			var twoCell = '<xdr:from><xdr:col>' + fromCol + '</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>' + fromRow + '</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>';
			twoCell += '<xdr:to><xdr:col>' + toCol + '</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>' + toRow + '</xdr:row><xdr:rowOff>99999</xdr:rowOff></xdr:to>';
			twoCell += '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="' + (i + 1) + '" name="' + image.name + '"></xdr:cNvPr><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>';
			twoCell += '<xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId' + (i + 1) + '"/>';
			twoCell += '<a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/>';
			o[o.length] = writextag('xdr:twoCellAnchor', twoCell, image.attrs);
		}
		else if (pos.type === 'oneCellAnchor') {
			var from = pos.from || {}, ext = pos.ext || {},
				fromCol = from.col || 0, fromRow = from.row || 0,
				colOff = from.colOff || 0, rowOff = from.rowOff || 0,
				width = ext.width || 100000, height = ext.height || 100000;

			var oneCell = '<xdr:from><xdr:col>' + fromCol + '</xdr:col><xdr:colOff>' + colOff + '</xdr:colOff><xdr:row>' + fromRow + '</xdr:row><xdr:rowOff>' + rowOff + '</xdr:rowOff></xdr:from>';
			oneCell += '<xdr:ext cx="' + width + '" cy="' + height + '"/>';
			oneCell += '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="' + (i + 1) + '" name="' + image.name + '"></xdr:cNvPr><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>';
			oneCell += '<xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId' + (i + 1) + '"/>';
			oneCell += '<a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/>';
			o[o.length] = writextag('xdr:oneCellAnchor', oneCell, image.attrs);
		}
		else if (pos.type === 'absoluteAnchor') {
			var posXY = pos.pos || {}, ext = pos.ext || {},
				x = posXY.x || 0, y = posXY.y || 0,
				width = ext.width || 100000, height = ext.height || 100000;

			var absAnchor = '<xdr:pos x="' + x + '" y="' + y + '"/>';
			absAnchor += '<xdr:ext cx="' + width + '" cy="' + height + '"/>';
			absAnchor += '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="' + (i + 1) + '" name="' + image.name + '"></xdr:cNvPr><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>';
			absAnchor += '<xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId' + (i + 1) + '"/>';
			absAnchor += '<a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/>';
			o[o.length] = writextag('xdr:absoluteAnchor', absAnchor, image.attrs);
		}
	}

	if (o.length > 2) {
		o[o.length] = '</xdr:wsDr>';
		o[1] = o[1].replace("/>", ">");
	}
	return o.join("");
}

// Helper function to create XML tags (assumed to be available)
function writextag(tag, content, attrs) {
	var out = '<' + tag;
	if (attrs) {
		for (var key in attrs) {
			out += ' ' + key + '="' + attrs[key] + '"';
		}
	}
	return out + (content ? '>' + content + '</' + tag + '>' : '/>');
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
