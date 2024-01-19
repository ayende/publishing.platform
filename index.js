const fs = require('fs').promises;
const path = require('path');
const { authenticate } = require('@google-cloud/local-auth');
const { google } = require('googleapis');
const admZip = require("adm-zip");
const { default: jssoup } = require('jssoup');
const MetaWeblog = require('metaweblog-api');
const flourite = require('flourite');
const Prism = require('prismjs');
const { URL } = require('url');

const SCOPES = [
    "https://www.googleapis.com/auth/documents",
    "https://www.googleapis.com/auth/drive.readonly"
];


const TOKEN_PATH = 'C:/Work/Credentials/token.json';
const CREDENTIALS_PATH = 'C:/Work/Credentials/credentials.json';


async function loadSavedCredentialsIfExist() {
    try {
        const content = await fs.readFile(TOKEN_PATH);
        const credentials = JSON.parse(content);
        return google.auth.fromJSON(credentials);
    } catch (err) {
        return null;
    }
}

async function saveCredentials(client) {
    const content = await fs.readFile(CREDENTIALS_PATH);
    const keys = JSON.parse(content);
    const key = keys.installed || keys.web;
    const payload = JSON.stringify({
        type: 'authorized_user',
        client_id: key.client_id,
        client_secret: key.client_secret,
        refresh_token: client.credentials.refresh_token,
    });
    await fs.writeFile(TOKEN_PATH, payload);
}

async function authorize() {
    let client = await loadSavedCredentialsIfExist();
    if (client) {
        return client;
    }
    client = await authenticate({
        scopes: SCOPES,
        keyfilePath: CREDENTIALS_PATH,
    });
    if (client.credentials) {
        await saveCredentials(client);
    }
    return client;
}


async function processFile(auth, fileId) {
    const drive = google.drive({ version: 'v3', auth: auth });
    const file = await drive.files.export({
        fileId: fileId,
        mimeType: 'application/zip',
    }, { responseType: 'arraybuffer' });
    const text = await drive.files.export({
        fileId: fileId,
        mimeType: 'text/plain'
    });

    let blocks = [];
    for (const match of text.data.matchAll(/\uEC03(.*?)\uEC02/gs)) {
        const code = match[1].trim();
        const lang = flourite(code, { shiki: true, noUnkown: true }).language;
        const formattedCode = Prism.highlight(code, Prism.languages[lang], lang);

        blocks.push("<hr/><pre class='line-numbers language-" + lang + ">" +
            "<code class='line-numbers language-" + lang + "'>" +
            formattedCode + "</code></pre><hr/>");
    }

    const zip = new admZip(Buffer.from(file.data));
    const entries = zip.getEntries();
    var htmlText = entries.find(e => path.extname(e.entryName) == '.html')
        .getData().toString('utf8');

    const htmlDoc = new jssoup(htmlText).find('body');

    let tags = [];
    let postId = null;
    let codeSegmentIndex = 0;
    let inCodeSegment = false;
    htmlDoc.findAll().forEach(e => {
        var text = e.getText().trim();
        if (text == "&#60419;") {
            e.replaceWith(blocks[codeSegmentIndex++]);
            inCodeSegment = true;
        }
        if (inCodeSegment) {
            e.extract();
        }
        if (text == "&#60418;") {
            inCodeSegment = false;
        }
    })
    htmlDoc.findAll('span').forEach(span => {
        var text = span.getText();

        if (text.startsWith('Tags:')) {
            tags = text.substring("Tags:".length).split(',').map(t => t.trim());
            span.extract();
        }
        if (text.startsWith('PostId:')) {
            postId = text.split(':')[1].trim();
            span.extract();
        }
    });

    htmlDoc.findAll('a').forEach(a => {
        if (a.attrs.hasOwnProperty('href')) {
            // Google Docs put a redirect like: https://www.google.com/url?q=ACTUAL_URL
            var link = new URL(a.attrs.href);
            var q = link.searchParams.get('q');
            if (q) {
                a.attrs.href = q;
            }
        }
    });
    htmlDoc.findAll('img').forEach(img => {
        if (img.attrs.hasOwnProperty('src')) {
            let src = img.attrs.src;
            let imgName = src.split('/').pop();
            let imgData = entries.find(e => e.entryName === 'images/' + imgName).getData();
            let imgType = imgName.split('.').pop();
            let imgSrc = 'data:image/' + imgType + ';base64,' + imgData.toString('base64');
            img.replaceWith('<img src="' + imgSrc + '" style="float: right"/>');
        }
    })

    var currentElement = htmlDoc.nextElement;
    let cleanHTML = cleanHtml(currentElement);
    while (currentElement.nextSibling !== undefined) {
        currentElement = currentElement.nextSibling;
        cleanHTML += cleanHtml(currentElement);
    }

    cleanHTML += `\r\n<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/prism/9000.0.1/themes/prism.min.css" integrity="sha512-/mZ1FHPkg6EKcxo0fKXF51ak6Cr2ocgDi5ytaTBjsQZIH/RNs6GF6+oId/vPe3eJB836T36nXwVh/WBl/cWT4w==" crossorigin="anonymous" referrerpolicy="no-referrer" />`;

    return [cleanHTML, postId, tags];
}

function cleanHtml(htmlElement) {
    if (htmlElement.hasOwnProperty('contents') === false) {
        let text = htmlElement._text;
        if (text === "&nbsp;" && htmlElement.parent.name !== "td") {
            return '';
        }
        return htmlElement._text;
    }
    let modifiers = {
        bold: false,
        italics: false,
        styles: ''
    };
    let newStyles = '';
    if (htmlElement.attrs.hasOwnProperty('style')) {
        newStyles = computeNewStyles(htmlElement, newStyles, modifiers);
    }
    modifiers['styles'] = newStyles;
    let href;
    if (htmlElement.name === "a") {
        href = ' href="' + htmlElement.attrs.href + '"';
    } else {
        href = '';
    }
    let newElemStart;
    let newElemEnd;
    if (newStyles) {
        newElemStart = '<' + htmlElement.name + ' style="' + modifiers['styles'] + href + '">';
        newElemEnd = '</' + htmlElement.name + '>';
    } else {
        newElemStart = '<' + htmlElement.name + href + '>';
        newElemEnd = '</' + htmlElement.name + '>';
    }
    let contents = htmlElement.contents;
    let newContents = '';
    for (let i = 0; i < contents.length; i++) {
        newContents += cleanHtml(contents[i]);
    }
    if (newContents === "") {
        return '';
    }
    if (htmlElement.name === "span" && !modifiers['styles'] || htmlElement.name === "b") {
        newElemStart = '';
        newElemEnd = '';
    }
    let newHtmlElement = `${newElemStart}${modifiers['bold'] ? '<strong>' : ''}${modifiers['italics'] ? '<em>' : ''}${newContents}${modifiers['italics'] ? '</em>' : ''}${modifiers['bold'] ? '</strong>' : ''}${newElemEnd}`;
    if (htmlElement.name === "td" || htmlElement.name === "li") {
        newHtmlElement = newHtmlElement.replace(/<p.*?>/g, "");
        newHtmlElement = newHtmlElement.replace('</p>', "");
    }
    return newHtmlElement;
}


function computeValidStyles(htmlElement) {
    const headings = new Set(['h1', 'h2', 'h3', 'h4', 'h5']);
    let validStyles;
    if (htmlElement.parent.name === "a") {
        validStyles = ['color', 'font-style', 'font-weight', 'text-align'];
    } else if (headings.has(htmlElement.parent.name)) {
        validStyles = ['text-align'];
    } else if (htmlElement.name === "li" || htmlElement.name === "b") {
        validStyles = [];
    } else {
        validStyles = ['color', 'font-style', 'font-weight', 'text-decoration', 'text-decoration-line', 'text-align'];
    }
    return new Set(validStyles);
}

function computeNewStyles(htmlElement, modifiers) {
    const validStyles = computeValidStyles(htmlElement);

    let newStyles = '';
    for (let style of htmlElement.attrs.style.split(';')) {
        let parts = style.split(':');
        if (parts.length !== 2)
            continue;
        let styleType = parts[0].trim(), styleValue = parts[1].trim();

        if (validStyles.has(styleType) === false || styleValue === undefined)
            continue;

        if (styleType === 'color' && (styleValue !== "#000000" && styleValue !== "#1155cc")) {
            newStyles += 'color:' + styleValue + ';';
        }
        if (styleType === 'font-style' && styleValue === 'italic') {
            modifiers['italics'] = true;
        }
        if (styleType === 'font-weight' && styleValue > 500) {
            modifiers['bold'] = true;
        }
        if ((styleType === 'text-decoration-line' || styleType === 'text-decoration') && styleValue === 'underline') {
            newStyles += 'text-decoration:underline;';
        }
        if (styleType === 'text-align') {
            newStyles += 'text-align:' + styleValue + ';';
        }
    }
    return newStyles;
}

async function getBlogClient() {
    const content = await fs.readFile(CREDENTIALS_PATH);
    const blog = JSON.parse(content).blog;
    const blogApi = new MetaWeblog('https://ayende.com/blog/Services/MetaWeblogAPI.ashx');
    return {
        getPost: async (postId) => {
            return await blogApi.getPost(postId, blog.username, blog.password);
        },
        editPost: async (postId, post) => {
            return await blogApi.editPost(postId, blog.username, blog.password, post, true);
        },
        createPost: async (post) => {
            return await blogApi.newPost('', blog.username, blog.password, post, true);
        }
    }
}

(async () => {

    let arg = process.argv[2];

    try {
        const url = new URL(arg);
        const parts = url.pathname.split('/');
        arg = parts[parts.length - 2];
    } catch {
        arg = process.argv[1];
    }

    const file = arg;
    const auth = await authorize();
    const [html, postId, tags] = await processFile(auth, file);
    const docs = google.docs({ version: 'v1', auth });
    const doc = await docs.documents.get({ documentId: file });

    var post = {
        description: html,
        title: doc.data.title,
        categories: tags,
    };

    const blogClient = await getBlogClient();

    if (postId) {
        await blogClient.editPost(postId, post);
    }
    else {
        const newPostId = await blogClient.createPost(post);
        await docs.documents.batchUpdate({
            documentId: documentId,
            requestBody: {
                requests: [{
                    insertText: {
                        text: `PostId: ${newPostId}\n`,
                        location: { index: 1 }
                    }
                }]
            }
        });
    }
    console.log('Published: ' + doc.data.title);

})()