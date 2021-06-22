/**
 * @file
 * Custom Example: Remote signers, cc, envelope has three documents
 * @author Christopher Gitmed
 */

const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');

const path = require('path')
  , fs = require('fs-extra')
  , docusign = require('docusign-esign')
  , validator = require('validator')
  , dsConfig = require('../../config/index.js').config
  ;

const signatories = [
  {
    email: 'cgitmed+1@gmail.com',
    name: 'Bart'
  },
  {
    email: 'cgitmed+2@gmail.com',
    name: 'Lisa'
  },
  {
    email: 'cgitmed+3@gmail.com',
    name: 'Homer'
  },
  {
    email: 'cgitmed+4@gmail.com',
    name: 'Marge'
  }
]

const ccRecipients = [
  {
    email: 'chrisgitmed@yahoo.com',
    name: 'Krusty'
  }
]

/* PRE-PROCESSING THE FILE */

// The error object contains additional information when logged with JSON.stringify (it contains a properties object containing all suberrors).
function replaceErrors(key, value) {
  if (value instanceof Error) {
    return Object.getOwnPropertyNames(value).reduce(function (error, key) {
      error[key] = value[key];
      return error;
    }, {});
  }
  return value;
}

function errorHandler(error) {
  console.log(JSON.stringify({ error: error }, replaceErrors));

  if (error.properties && error.properties.errors instanceof Array) {
    const errorMessages = error.properties.errors.map(function (error) {
      return error.properties.explanation;
    }).join("\n");
    console.log('errorMessages', errorMessages);
    // errorMessages is a humanly readable message looking like this :
    // 'The tag beginning with "foobar" is unopened'
  }
  throw error;
}

// CREATE THE FILE

// Load the docx file as binary content
var content = fs
  .readFileSync(path.resolve(__dirname, '../../demo_documents/test.docx'), 'binary');

const zip = new PizZip(content);
let doc;
try {
  doc = new Docxtemplater(zip);
} catch (error) {
  // Catch compilation errors (errors caused by the compilation of the template : misplaced tags)
  errorHandler(error);
}

//set the templateVariables
let signatureString = '';
for (let i = 0; i < signatories.length; i++) {
  signatureString += `/sn${i + 2}/ `
}
doc.setData({
  template_string: signatureString
});

try {
  // render the document (replace all occurences of {template_string} by signatureString)
  doc.render()
}
catch (error) {
  // Catch rendering errors (errors relating to the rendering of the template : angularParser throws an error)
  errorHandler(error);
}

var buf = doc.getZip()
  .generate({ type: 'nodebuffer' });

// buf is a nodejs buffer, you can either write it to a file or do anything else with it.
fs.writeFileSync(path.resolve(__dirname, '../../demo_documents/test2.docx'), buf);


const CustomSigningViaEmail = exports
  , eg = 'eg002' // This example reference.
  , mustAuthenticate = '/ds/mustAuthenticate'
  , minimumBufferMin = 3
  , demoDocsPath = path.resolve(__dirname, '../../demo_documents')
  , doc2File = 'test2.docx'
  , doc3File = 'World_Wide_Corp_lorem.pdf'
  ;

/**
 * Create the envelope
 * @param {object} req Request obj
 * @param {object} res Response obj
 */
CustomSigningViaEmail.createController = async (req, res) => {
  // Step 1. Check the token
  // At this point we should have a good token. But we
  // double-check here to enable a better UX to the user.
  let tokenOK = req.dsAuth.checkToken(minimumBufferMin);
  if (!tokenOK) {
    req.flash('info', 'Sorry, you need to re-authenticate.');
    // Save the current operation so it will be resumed after authentication
    req.dsAuth.setEg(req, eg);
    res.redirect(mustAuthenticate);
  }

  // Step 2. Call the worker method
  let body = req.body
    // Additional data validation might also be appropriate
    , signerEmail = validator.escape(body.signerEmail)
    , signerName = validator.escape(body.signerName)
    , ccEmail = validator.escape(body.ccEmail)
    , ccName = validator.escape(body.ccName)
    , envelopeArgs = {
      signerEmail: signerEmail,
      signerName: signerName,
      ccEmail: ccEmail,
      ccName: ccName,
      status: "sent"
    }
    , args = {
      accessToken: req.user.accessToken,
      basePath: req.session.basePath,
      accountId: req.session.accountId,
      envelopeArgs: envelopeArgs
    }
    , results = null
    ;

  try {
    results = await CustomSigningViaEmail.worker(args)
  }
  catch (error) {
    let errorBody = error && error.response && error.response.body
      // we can pull the DocuSign error code and message from the response body
      , errorCode = errorBody && errorBody.errorCode
      , errorMessage = errorBody && errorBody.message
      ;
    // In production, may want to provide customized error messages and
    // remediation advice to the user.
    res.render('pages/error', { err: error, errorCode: errorCode, errorMessage: errorMessage });
  }
  if (results) {
    req.session.envelopeId = results.envelopeId; // Save for use by other examples
    // which need an envelopeId
    res.render('pages/example_done', {
      title: "Envelope sent",
      h1: "Envelope sent",
      message: `The envelope has been created and sent!<br/>Envelope ID ${results.envelopeId}.`
    });
  }
}

/**
 * This function does the work of creating the envelope
 */
// ***DS.snippet.0.start
CustomSigningViaEmail.worker = async (args) => {
  // Data for this method
  // args.basePath
  // args.accessToken
  // args.accountId

  let dsApiClient = new docusign.ApiClient();
  dsApiClient.setBasePath(args.basePath);
  dsApiClient.addDefaultHeader('Authorization', 'Bearer ' + args.accessToken);
  let envelopesApi = new docusign.EnvelopesApi(dsApiClient)
    , results = null;

  // Step 1. Make the envelope request body
  let envelope = makeEnvelope(args.envelopeArgs)

  // Step 2. call Envelopes::create API method
  // Exceptions will be caught by the calling function
  results = await envelopesApi.createEnvelope(args.accountId,
    { envelopeDefinition: envelope });
  let envelopeId = results.envelopeId;

  console.log(`Envelope was created. EnvelopeId ${envelopeId}`);
  return ({ envelopeId: envelopeId })
}

/**
 * Creates envelope
 * @function
 * @param {Object} args parameters for the envelope
 * @returns {Envelope} An envelope definition
 * @private
 */
function makeEnvelope(args) {
  // Data for this method
  // args.signerEmail
  // args.signerName
  // args.ccEmail
  // args.ccName
  // args.status
  // demoDocsPath (module constant)
  // doc2File (module constant)
  // doc3File (module constant)

  let doc2DocxBytes, doc3PdfBytes;
  // read files from a local directory
  // The reads could raise an exception if the file is not available!
  doc2DocxBytes = fs.readFileSync(path.resolve(demoDocsPath, doc2File));
  doc3PdfBytes = fs.readFileSync(path.resolve(demoDocsPath, doc3File));

  // create the envelope definition
  let env = new docusign.EnvelopeDefinition();
  env.emailSubject = 'Please sign this document set';

  // add the documents
  let doc1 = new docusign.Document()
    , doc1b64 = Buffer.from(document1(args)).toString('base64')
    , doc2b64 = Buffer.from(doc2DocxBytes).toString('base64')
    , doc3b64 = Buffer.from(doc3PdfBytes).toString('base64')
    ;

  doc1.documentBase64 = doc1b64;
  doc1.name = 'Order acknowledgement'; // can be different from actual file name
  doc1.fileExtension = 'html'; // Source data format. Signed docs are always pdf.
  doc1.documentId = '1'; // a label used to reference the doc

  // Alternate pattern: using constructors for docs 2 and 3...
  let doc2 = new docusign.Document.constructFromObject({
    documentBase64: doc2b64,
    name: 'Battle Plan', // can be different from actual file name
    fileExtension: 'docx',
    documentId: '2'
  });

  let doc3 = new docusign.Document.constructFromObject({
    documentBase64: doc3b64,
    name: 'Lorem Ipsum', // can be different from actual file name
    fileExtension: 'pdf',
    documentId: '3'
  });

  // The order in the docs array determines the order in the envelope
  env.documents = [doc1, doc2, doc3];

  // create a signer recipient to sign the document, identified by name and email
  // We're setting the parameters via the object constructor

  const signers = [];
  // GET THE FIRST SIGNER FROM ARGS
  let signer1 = docusign.Signer.constructFromObject({
    email: args.signerEmail,
    name: args.signerName,
    recipientId: '1',
    routingOrder: '1'
  });
  signers.push(signer1);

  // GET THE REST FROM THE SIGNATORIES ARRAY
  for(let i = 0; i < signatories.length; i++) {
    signatories[i].recipientId = `${i + 2}`; // Assign recipientID, starting from 2 to factor in the signer1
    signatories[i].routingOrder = '2' // Setting all routingOrders the same (for now)
    signers.push(signatories[i]);
  }

  // GET CARBON COPY RECIPIENTS
  for(let i = 0; i < ccRecipients.length; i++) {
    ccRecipients[i].recipientId = `${i + 2 + signatories.length}`;
    ccRecipients[i].routingOrder = '3';
  }

  // CREATE SIGNHERE FIELDS
  const signHereArr = [];
  for(let i = 0; i < signers.length; i++) {
    const signHere = docusign.SignHere.constructFromObject({
      anchorString: `/sn${i+1}/`,
      anchorYOffset: '10', anchorUnits: 'pixels',
      anchorXOffset: '20'
    })
    signHereArr.push(signHere)
  }

  // LOOP THROUGH SIGNERS AND ASSIGN TABS
  for(let i = 0; i < signers.length; i++) {
    const signerTabs = docusign.Tabs.constructFromObject({
      signHereTabs: [signHereArr[i]]
    });
    signers[i].tabs = signerTabs;
  }

  let recipients = docusign.Recipients.constructFromObject({
    signers: signers,
    carbonCopies: ccRecipients
  });

  console.log('recipients: ', recipients);

  env.recipients = recipients;

  // Request that the envelope be sent by setting |status| to "sent".
  // To request that the envelope be created as a draft, set to "created"
  env.status = args.status;

  return env;
}

/**
 * Creates document 1
 * @function
 * @private
 * @param {Object} args parameters for the envelope
 * @returns {string} A document in HTML format
 */

function document1(args) {
  // Data for this method
  // args.signerEmail
  // args.signerName
  // args.ccEmail
  // args.ccName

  return `
    <!DOCTYPE html>
    <html>
        <head>
          <meta charset="UTF-8">
        </head>
        <body style="font-family:sans-serif;margin-left:2em;">
        <h1 style="font-family: 'Trebuchet MS', Helvetica, sans-serif;
            color: darkblue;margin-bottom: 0;">World Wide Corp</h1>
        <h2 style="font-family: 'Trebuchet MS', Helvetica, sans-serif;
          margin-top: 0px;margin-bottom: 3.5em;font-size: 1em;
          color: darkblue;">Order Processing Division</h2>
        <h4>Ordered by ${args.signerName}</h4>
        <p style="margin-top:0em; margin-bottom:0em;">Email: ${args.signerEmail}</p>
        <p style="margin-top:0em; margin-bottom:0em;">Copy to: ${args.ccName}, ${args.ccEmail}</p>
        <p style="margin-top:3em;">
  Candy bonbon pastry jujubes lollipop wafer biscuit biscuit. Topping brownie sesame snaps sweet roll pie. Croissant danish biscuit soufflé caramels jujubes jelly. Dragée danish caramels lemon drops dragée. Gummi bears cupcake biscuit tiramisu sugar plum pastry. Dragée gummies applicake pudding liquorice. Donut jujubes oat cake jelly-o. Dessert bear claw chocolate cake gummies lollipop sugar plum ice cream gummies cheesecake.
        </p>
        <!-- Note the anchor tag for the signature field is in white. -->
        <h3 style="margin-top:3em;">Agreed: <span style="color:white;">**signature_1**/</span></h3>
        </body>
    </html>
  `
}
// ***DS.snippet.0.end

/**
 * Form page for this application
 */
CustomSigningViaEmail.getController = (req, res) => {
  // Check that the authentication token is ok with a long buffer time.
  // If needed, now is the best time to ask the user to authenticate
  // since they have not yet entered any information into the form.
  let tokenOK = req.dsAuth.checkToken();
  if (tokenOK) {
    res.render('pages/examples/eg002SigningViaEmail.ejs', {
      eg: eg, csrfToken: req.csrfToken(),
      title: "Signing request by email",
      sourceFile: path.basename(__filename),
      sourceUrl: dsConfig.githubExampleUrl + 'eSignature/' + path.basename(__filename),
      documentation: dsConfig.documentation + eg,
      showDoc: dsConfig.documentation
    });
  } else {
    // Save the current operation so it will be resumed after authentication
    req.dsAuth.setEg(req, eg);
    res.redirect(mustAuthenticate);
  }
}
