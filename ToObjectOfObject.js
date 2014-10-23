importClass(Packages.com.sun.star.uno.UnoRuntime);
importClass(Packages.com.sun.star.frame.XModel);
importClass(Packages.com.sun.star.sheet.XSpreadsheetDocument);
importClass(Packages.com.sun.star.sheet.XSpreadsheetView);
importClass(Packages.com.sun.star.container.XNamed);

function isNumber(n) {

  return !isNaN(parseFloat(n)) && isFinite(n);
  
}

function toArray(sheet) {
	
	var master = [];
	
	var rows = 0;
	var columns = 0;
	var object = sheet.getObject();
	
	while (true) {
		var cell = object.getCellByPosition(0, rows);
		var content = cell.getFormula();
		if (content == "") break;
		rows ++;
	}
	
	while (true) {
		var cell = object.getCellByPosition(columns, 0);
		var content = cell.getFormula();
		if (content == "") break;
		columns ++;
	}
	
	for (var j=0; j<rows; j++) {
		var arr = [];
		for (var i=0; i<columns; i++) {
			var cell = object.getCellByPosition(i, j);
			var content = cell.getFormula();
			var text = content + "";
			var value = cell.getValue();
			if (isNumber(text)) {
				arr.push(value);
				continue;
			}
			if (content.indexOf(",") != -1) {
				arr.push("\"" + text + "\"");
			}else {
				arr.push(text);
			}
		}
		master.push(arr);
	}
	
	return master;
	
}

function toObjectOfObject(sheet) {
	
	var master = {};
	var array = toArray(sheet);
	var keys = [];
	
	for (var i in array) {
		var row = array[i];
		if (i == 0) {
			keys = row;
			// make all keys lowercase and no spaces
			for (var j in keys) {
				var origin = keys[j];
				var dest = origin.replace(/ /g, "_");
				dest = dest.toLowerCase();
				keys[j] = dest;
			}
			continue;
		}
		var obj = {};
		if (keys.length > 2) {
			for (var j in keys) {
				if (j == 0) continue;
				obj[keys[j]] = row[j];
			}
		}else {
			obj = row[1];
		}
		master[row[0]] = obj;
	}
	
	return master;
	
}

// a Swing UI for displaying the data
function EditorPane( ) {

	Swing = Packages.javax.swing;
	this.pane = new Swing.JEditorPane("text/html","" );
	this.jframe = new Swing.JFrame( );
	this.jframe.setBounds( 100,100,500,400 );
	var editorScrollPane = new Swing.JScrollPane(this.pane);
	editorScrollPane.setVerticalScrollBarPolicy(
	Swing.JScrollPane.VERTICAL_SCROLLBAR_ALWAYS);
	editorScrollPane.setPreferredSize(new java.awt.Dimension(150, 150));
	editorScrollPane.setMinimumSize(new java.awt.Dimension(10, 10));
	this.jframe.setVisible( true );
	this.jframe.getContentPane().add( editorScrollPane );
	
	// public methods
	this.getPane = function( ) { return this.pane; }
	this.getJFrame = function( ) { return this.jframe; }

}

( function main( ) {

	//get the document object from the scripting context
	oDoc = XSCRIPTCONTEXT.getDocument();
	
	//get the XSpreadsheetDocument interface from the document
	xSpreadsheetDocument = UnoRuntime.queryInterface(XSpreadsheetDocument, oDoc);

	// get a reference to the sheets for this doc
	var xDocModel = UnoRuntime.queryInterface(XModel, xSpreadsheetDocument);
	var xSpreadsheetModel = UnoRuntime.queryInterface(XModel, xDocModel);
	var xSpreadsheetController = xSpreadsheetModel.getCurrentController();
	var xSpreadsheetView = UnoRuntime.queryInterface(XSpreadsheetView, xSpreadsheetController);
	var activeSheet = xSpreadsheetView.getActiveSheet();
	var xSheet = UnoRuntime.queryInterface(XNamed, activeSheet);

	var sheets = xSpreadsheetDocument.getSheets();
	var sheet = sheets.getByName(xSheet.getName());

	// construct a new EditorPane
	var editor = new EditorPane( );
	var pane = editor.getPane( );
	
	// compose
	var data =  toObjectOfObject(sheet);
	var text =  JSON.stringify(data);
	
	pane.setText(text);

})();
