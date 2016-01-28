/* DumpIndexToExcel commander component
 * To use add require('../cmds/dumpindextoexcel.js')(program) to your commander.js based node executable before program.parse
 */
'use strict';


//This exposes the command for use by autocmd.
module.exports = DumpIndexToExcel;

// Load any required modules
var elasticsearch = require('elasticsearch');
var _ = require('lodash');
var path = require('path');
var XLSX = require('xlsx');

//Variable to hold internal items shared by all instances
var internals = {};

/**
 * Dumps an index to an Excel spreadsheet
 * @param {[type]} program The commander program
 */
function DumpIndexToExcel(program) {

	program
		.command('dumpindextoexcel <indexname> <outputfile>')
		.version('0.0.0')		
		.description('Exports the URLs and some metadata contained in an index to an Excel Spreadsheet.  indexname should be the name of the index, and outputfile is the XLSX file to create.')
		.option(
			'-r, --reportfields <fields>',
			'A comma separated list of fields to extract for the report.',
			function(val){
				//Extract fields, this should be a comma separated list.
				return String.split(val);
			},
			['host', 'url', 'type', 'contentLength', 'title']
		)
		.option(
			'-h, --hostfilter <hostname>',
			'A hostname to filter the items for.  (e.g. www.cancer.gov)'
		)
		.action(internals.commandAction);	
};

/**
 * Command action from handing this command.
 * @param  {[type]} indexname  the index to scrape
 * @param  {[type]} outputfile the excel file to write the outputs to
 * @param  {[type]} cmd        the command object
 * @return {[type]}            [description]
 */
internals.commandAction = function(indexname, outputfile, cmd) {

	//Check if the output file exists or not, and if it has .xlsx at the end or not.
	var cleanpath = internals.validateAndCleanFileName(outputfile);

	if (cmd.parent.verbose) {
		console.log('Index Name: ' + indexname);
		console.log('Report Fields: ' + cmd.reportfields);
		console.log('Server: ' + cmd.parent.server);
		console.log('Port: ' + cmd.parent.port);
		console.log('Host filter: ' + cmd.hostfilter);
	}

	//Now let's get the results!
	internals.fetchResults(
		{
			indexname: indexname, 
			reportfields: cmd.reportfields, 
			server: cmd.parent.server,
			port: cmd.parent.port,
			hostfilter: cmd.hostfilter
		}, 
		function(err, results) {

			if (err) {
				//process.stderr.write(err);
				console.log(err)
				process.exit(10);
			}

			internals.outputExcel(
				{
					outputfile: cleanpath,
					reportfields: cmd.reportfields
				},
				results,
				function(outputErr) {

					if (outputErr) {
						console.log(outputErr)
						process.exit(20);	
					}
					
					//Eh, dunno, do something
					console.log('Results written to ' + cleanpath);

					process.exit(0);					
				}
			);
		}
	)	
}

/**
 * [outputExcel description]
 * @param  {[type]} params [description]
 * @param  {[type]} params.outputfile The file to write results to.
 * @param  {[type]} params.reportfields A list of fields for reporting.
 * @param  {[type]} items  the items to write out.
 * * @param  {[type]} completion  a callback to use when the results are done.
 * @return {[type]}        [description]
 */
internals.outputExcel = function(params, items, completion) {

	//TODO: validate params

	//Setup workbook and sheet	
	var workbook = new internals.Workbook();
	var sheet = workbook.addSheet('indexitems');

	//We will use this in multiple places, so let's set it here.
	var field_count = params.reportfields.length;
	var item_count = items.length;

	//Add header row using the reportfields at the headers
	for (var head_col=0; head_col < field_count; head_col++) {
		var cell_addr = XLSX.utils.encode_cell({c:head_col,r:0});

		sheet[cell_addr] = {
			v: params.reportfields[head_col], // Value of Cell
			t: 's' // Type of cell
		}
	}	

	//Now, loop through the results
	for (var item_idx = 0; item_idx < item_count; item_idx++) {

		//Loop through each field that on the items.
		for (var field_col = 0; field_col < field_count; field_col++) {
			sheet[workbook.getCellID(field_col, item_idx)] = {
				v: items[item_idx][params.reportfields[field_col]], //Value
				t: 's' // Set type as string.  It would be cool in the future to set the correct type.
			}
		}
	}

	//Now we have to setup the worksheet ranges
	//Set the range
	sheet['!ref'] = XLSX.utils.encode_range({
		s: { //First Cell in the range
			c: 0, 
			r: 0
		}, 
		e: { //Last Cell in the range
			c: field_count, 
			r: item_count + 1
		}
	});	
	
	//Output the XLSX
	try {
		XLSX.writeFile(workbook, params.outputfile);
	} catch (ex) {
		completion(ex);		
	}

	completion(false);
}


/**
 * Fetch the items from the ElasticSearch server.
 * @param  {[type]} params    the options for the search
 * @param  {[type]} params.indexname    the index to scrape
 * @param  {[type]} params.reportfields the fields to report on
 * @param  {[type]} params.server       the server name
 * @param  {[type]} params.port         the server port
 * @param  {[type]} rescallback  a callback function for dealing with the results.
 * @return {[type]}              [description]
 */
internals.fetchResults = function(params, rescallback) {

	var items = [];


	// This is referenced too often to keep typing params.reportFields.
	var reportFields = params.reportfields;

	//Setup elastic search client.
	var client = new elasticsearch.Client({
		host: 'http://' + params.server + ':' + params.port,
		log: 'error'
	});

	//Setup search options as separate variable so I can modify things
	//like the query if it is set, otherwise, I do not need to pass it in
	//to the search function.
	var searchOpts = {
		index: params.indexname,
		scroll: '1s',
		fields: reportFields
	};

	if (params.hostfilter) {
		searchOpts.q = 'host:' + params.hostfilter
	}


	client.search(
		searchOpts
		, function getMoreUntilDone(error, response) {

		//Encountered search error, bail.
		if (error) {
			rescallback(error, false)
		}

		//loop through results.
		response.hits.hits.forEach(function (hit) {
		    if (hit.fields != null) {

	    		var res = {};

	    		for (var i = 0; i< reportFields.length; i++) {
	    			res[reportFields[i]] = hit.fields[reportFields[i]];
	    		}

				//Push result into list
				items.push(res);
			} else {
				items.push({});
				//console.log("Null Result");
			}
		});

		//more to fetch, well then "scroll" to next page of results
		if (response.hits.total !== items.length) {
			client.scroll({
				scrollId: response._scroll_id,
				scroll: '1s'
			}, getMoreUntilDone);
		} else {
			rescallback(false, items);
		}
	})

}

/**
 * Validates the file name an cleans stuff up.
 * @param  {[type]} outputfile the file path as passed to the program
 * @return {[type]}            the cleaned up path
 */
internals.validateAndCleanFileName = function(outputfile) {

	//First, get a valid file name from this.  If there is no extension, slap on a .xlsx.	
	//Handle things like .., ., //, etc.
	var cleanpath = path.normalize(outputfile);

	//Figure out the full path for this item
	cleanpath = path.resolve(cleanpath);

	var ext = path.extname(cleanpath);

	//If there is an extension, then it better be .xlsx
	if (ext && ext.toLowerCase() !== '.xlsx') {
		process.stderr.write('Error: Extension, ' + ext + ', not allowed.  Must end in xlsx!\n');
		process.exit(5);
	} else if (!ext) {
		cleanpath += '.xlsx';
	}
	
	return cleanpath;
}

////////////////////////////////////////////////////////////////////////////
///
///  Wrapper around XLSX stuff to create a nice and easy to use Workbook object.
///
////////////////////////////////////////////////////////////////////////////

/**
 * Object describing an Excel workbook
 */
internals.Workbook = function() {
	if(!(this instanceof internals.Workbook)) return new internals.Workbook();
	this.SheetNames = [];
	this.Sheets = {};
}

/**
 * Gets an encoded CellID for a column and row, accounting for a header row
 * @param  {[type]} col The column index
 * @param  {[type]} row The row index
 * @return {[type]}     The cell id
 */
internals.Workbook.prototype.getCellID = function(col, row) {
	return XLSX.utils.encode_cell({c:col, r: (row + 1)});	
}

/**
 * Adds a sheet to this workbook and returns the sheet
 * @param {[type]} sheetName [description]
 * @return {[type]}     The newly added sheet
 */
internals.Workbook.prototype.addSheet = function(sheetName) {

	var sheet = {};
	this.SheetNames.push(sheetName);
	this.Sheets[sheetName] = sheet;

	return sheet;
}
