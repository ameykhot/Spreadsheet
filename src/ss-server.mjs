import Path from 'path';

import express from 'express';
import bodyParser from 'body-parser';
import querystring from 'querystring';
import Mustache from './mustache.mjs';
import {AppError, Spreadsheet} from 'cs544-ss';
const STATIC_DIR = 'statics';
const TEMPLATES_DIR = 'templates';

//some common HTTP status codes; not all codes may be necessary
const OK = 200;
const CREATED = 201;
const NO_CONTENT = 204;
const BAD_REQUEST = 400;
const NOT_FOUND = 404;
const CONFLICT = 409;
const SERVER_ERROR = 500;

const __dirname = Path.dirname(new URL(import.meta.url).pathname);

export default function serve(port, store) {
  process.chdir(__dirname);
  const app = express();
  app.locals.port = port;
  app.locals.store = store;
  app.locals.mustache = new Mustache();
  app.use('/', express.static(STATIC_DIR));
  setupRoutes(app);
  app.listen(port, function() {
    console.log(`listening on port ${port}`);
  });
}


/*********************** Routes and Handlers ***************************/

function setupRoutes(app) {
  app.use(bodyParser.urlencoded({extended: true}));
  app.use(bodyParser.json())
  //new routes as per project 3 requirements
  app.get(`/`,renderOpenHtmlForm(app))
  app.post(`/`,postRequest(app))
  app.get(`/:ssName`,renderUpdateHtmlForm(app))
  app.post(`/:ssName`,updateFormRequest(app))


  app.use(do404(app));
  app.use(doErrors(app));

}

//@TODO add handlers
/****************************** Handlers *******************************/


// get api to render spreadsheet open page
function renderOpenHtmlForm(app){
  return (async function(req,res){
    try{
      const model = {msg : ''};
      res.send(app.locals.mustache.render('spreadsheetOpenPage',model));
    }
    catch (err){
      const mapped = mapError(err);
      res.status(mapped.status).json(mapped);
    }
  })
}

function postRequest(app){
  return (async function(req,res){
    try {
      let error = {};
      //validate ssName then redirected to post method
      if(validateField('ssName', req.body, error)) {
        res.redirect(`/${req.body.ssName}`);
      }
      else
      {
        const model = {msg : error.ssName};
        res.send(app.locals.mustache.render('spreadsheetOpenPage', model));
      }
    }
    catch (e){}
  });
}

function renderUpdateHtmlForm(app){
  return (async function(req,res){
      try {
        const spreadSheetName = req.params.ssName
        //created sheet of Spreadsheet class for query function
        let sheet = await Spreadsheet.make(req.params.ssName,app.locals.store)
        let result = {};
        result["ssName"] = spreadSheetName

        // read contents of spreadsheet
        result["data"] = await app.locals.store.readFormulas(spreadSheetName)

        //pass spreadsheet content to create view for mustache template
        result["table"] = getViewForTable(result,sheet)
        res.send(app.locals.mustache.render('spreadsheetUpdatePage', result));
      }
      catch (err){
        const mapped = mapError(err);
        res.status(mapped.status).json(mapped);
      }
  })
}

function getViewForTable(cellIDs,sheet) {
  let ttd = []
  let ttr = []
  let cell = ''
  let formula = []

  //check if spreadsheet is empty then create empty view
  if(cellIDs.data.length === 0){
    for (let i = 0; i <= MIN_ROWS ; i++)
    {
      formula = []
      for (let j = 0; j < MIN_COLS; j++)
      {
        cell = String.fromCharCode((97) + j) + i;
        //push empty data in spreadsheet td
        formula.push('')
        ttd[j] = {id: String.fromCharCode(97 + j).toUpperCase()}
        ttr[i] = {value: [i],cellVal :formula};
      }
    }
  }
  else {


    let cId = cellIDs.data.map(x => x[0]);
    //get value for those cellIds who have formula
    let forms = cellIDs.data.map(x => sheet.query(x[0]).value);

    //remove max element from row and col to iterate view accordingly
    let maxCol = cellIDs.data.map(x => x[0].charAt(0)).sort().slice(-1)
    let maxRow = cellIDs.data.map(x => x[0].substring(1)).sort((p, q) => p - q).slice(-1)
    let c = maxCol[0].toString().charCodeAt(0) - 96

    for (let i = 0; i <= (maxRow[0] > MIN_ROWS ? maxRow[0] : MIN_ROWS); i++) {
      formula = []
      for (let j = 0; j < (c > MIN_COLS ? c : MIN_COLS); j++) {
        cell = String.fromCharCode((97) + j) + i;
        formula.push(forms[cId.indexOf(cell)] === undefined ? '' : forms[cId.indexOf(cell)])
        ttd[j] = {id: String.fromCharCode(97 + j).toUpperCase()}
        ttr[i] = {value: [i], cellVal: formula};
      }
    }
  }
  //removed first zero element
  ttr.shift()
  return {ttr:ttr,ttd:ttd};
}

function updateFormRequest(app){
  return (async function(req,res){
      try{
        const spreadSheetName = req.params.ssName
        //Created sheet object to implement copy functionality
        let sheet = await Spreadsheet.make(spreadSheetName,app.locals.store)
        const cellID = req.body.cellId
        const formula = req.body.formula
        let error = {}

        //check for errors
        if(validateUpdate(req.body,error)) {
          //depending upon ssAct called the respective action
          switch (req.body.ssAct) {
            case 'updateCell' :
              await app.locals.store.updateCell(spreadSheetName, cellID, formula)
              break;
            case 'deleteCell':
              await app.locals.store.delete(spreadSheetName, cellID)
              break;
            case 'copyCell' :
              await sheet.copy(cellID, formula)
              break;
            case 'clear':
              await app.locals.store.clear(spreadSheetName)
              break;
          }
          res.redirect(`/${spreadSheetName}`);
        }
        else
        {
          //if error occurs store in result object then displayed them
          const spreadSheetName = req.params.ssName
          let sheet = await Spreadsheet.make(spreadSheetName,app.locals.store)
          let result = {};
          result["ssName"] = spreadSheetName
          result["data"] = await app.locals.store.readFormulas(spreadSheetName)
          result["ssActError"] = error.ssAct
          result["cellIdError"] = error.cellId
          result["formulaError"] = error.formula
          result["table"] = getViewForTable(result,sheet)
          res.send(app.locals.mustache.render('spreadsheetUpdatePage', result));

        }
      }
      catch (err){
        const mapped = mapError(err);
        res.status(mapped.status).json(mapped);
      }
  })
}



/** Default handler for when there is no route for a particular method
 *  and path.
 */
function do404(app) {
  return async function(req, res) {
    const message = `${req.method} not supported for ${req.originalUrl}`;
    res.status(NOT_FOUND).
      send(app.locals.mustache.render('errors', { errors: [{ msg: message, }] }));
  };
}

/** Ensures a server error results in an error page sent back to
 *  client with details logged on console.
 */ 
function doErrors(app) {
  return async function(err, req, res, next) {
    res.status(SERVER_ERROR);
    res.send(app.locals.mustache.render('errors',
					{ errors: [ {msg: err.message, }] }));
    console.error(err);
  };
}
/*************************** Mapping Errors ****************************/

const ERROR_MAP = {
}

/** Map domain/internal errors into suitable HTTP errors.  Return'd
 *  object will have a "status" property corresponding to HTTP status
 *  code and an error property containing an object with with code and
 *  message properties.
 */
function mapError(err) {
  const isDomainError = (err instanceof AppError);
  const status =
      isDomainError ? (ERROR_MAP[err.code] || BAD_REQUEST) : SERVER_ERROR;
  const error =
      isDomainError
          ? { code: err.code, message: err.message }
          : { code: 'SERVER_ERROR', message: err.toString() };
  if (!isDomainError) console.error(err);
  return { status, error };
}

/************************* SS View Generation **************************/

const MIN_ROWS = 10;
const MIN_COLS = 10;

//
//@TODO add functions to build a spreadsheet view suitable for mustache

/**************************** Validation ********************************/


const ACTS = new Set(['clear', 'deleteCell', 'updateCell', 'copyCell']);
const ACTS_ERROR = `Action must be one of ${Array.from(ACTS).join(', ')}.`;

//mapping from widget names to info.
const FIELD_INFOS = {
  ssAct: {
    friendlyName: 'Action',
    err: val => !ACTS.has(val) && ACTS_ERROR,
  },
  ssName: {
    friendlyName: 'Spreadsheet Name',
    err: val => !/^[\w\- ]+$/.test(val) && `
      Bad spreadsheet name "${val}": must contain only alphanumeric
      characters, underscore, hyphen or space.
    `,
  },
  cellId: {
    friendlyName: 'Cell ID',
    err: val => !/^[a-z]\d\d?$/i.test(val) && `
      Bad cell id "${val}": must consist of a letter followed by one
      or two digits.
    `,
  },
  formula: {
    friendlyName: 'cell formula',
  },
};

/** return true iff params[name] is valid; if not, add suitable error
 *  message as errors[name].
 */
function validateField(name, params, errors) {
  const info = FIELD_INFOS[name];
  const value = params[name];
  if (isEmpty(value)) {
    errors[name] = `The ${info.friendlyName} field must be specified`;
    return false;
  }
  if (info.err) {
    const err = info.err(value);
    if (err) {
      errors[name] = err;
      return false;
    }
  }
  return true;
}

  
/** validate widgets in update object, returning true iff all valid.
 *  Add suitable error messages to errors object.
 */
function validateUpdate(update, errors) {
  const act = update.ssAct ?? '';
  switch (act) {
    case '':
      errors.ssAct = 'Action must be specified.';
      return false;
    case 'clear':
      return validateFields('Clear', [], ['cellId', 'formula'], update, errors);
    case 'deleteCell':
      return validateFields('Delete Cell', ['cellId'], ['formula'],
			    update, errors);
    case 'copyCell': {
      const isOk = validateFields('Copy Cell', ['cellId','formula'], [],
				  update, errors);
      if (!isOk) {
	return false;
      }
      else if (!FIELD_INFOS.cellId.err(update.formula)) {
	  return true;
      }
      else {
	errors.formula = `Copy requires formula to specify a cell ID`;
	return false;
      }
    }
    case 'updateCell':
      return validateFields('Update Cell', ['cellId','formula'], [],
			    update, errors);
    default:
      errors.ssAct = `Invalid action "${act}`;
      return false;
  }
}

function validateFields(act, required, forbidden, params, errors) {
  for (const name of forbidden) {
    if (params[name]) {
      errors[name] = `
	${FIELD_INFOS[name].friendlyName} must not be specified
        for ${act} action
      `;
    }
  }
  for (const name of required) validateField(name, params, errors);
  return Object.keys(errors).length === 0;
}


/************************ General Utilities ****************************/

/** return new object just like paramsObj except that all values are
 *  trim()'d.
 */
function trimValues(paramsObj) {
  const trimmedPairs = Object.entries(paramsObj).
    map(([k, v]) => [k, v.toString().trim()]);
  return Object.fromEntries(trimmedPairs);
}

function isEmpty(v) {
  return (v === undefined) || v === null ||
    (typeof v === 'string' && v.trim().length === 0);
}

/** Return original URL for req.  If index specified, then set it as
 *  _index query param 
 */
function requestUrl(req, index) {
  const port = req.app.locals.port;
  let url = `${req.protocol}://${req.hostname}:${port}${req.originalUrl}`;
  if (index !== undefined) {
    if (url.match(/_index=\d+/)) {
      url = url.replace(/_index=\d+/, `_index=${index}`);
    }
    else {
      url += url.indexOf('?') < 0 ? '?' : '&';
      url += `_index=${index}`;
    }
  }
  return url;
}

