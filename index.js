const fs = require('fs');
const path = require('path');
const css = require('css');
const admZip = require("adm-zip");
const { default: jssoup } = require('jssoup');
const MetaWeblog = require('metaweblog-api');
const flourite = require('flourite');
const Prism = require('prismjs');
const { URL } = require('url');
const { exec } = require('child_process');

const loadLanguages = require('prismjs/components/');
const { cwd } = require('process');
loadLanguages(['lua', 'powershell', 'typescript', 'csharp',
    'fsharp', 'sql', 'bash', 'yaml', 'json', 'xml', 'markdown',
    'docker', 'ini', 'java', 'javascript', 'python', 'rust', 'swift', 'go',
    'ruby', 'php', 'perl', 'powershell', 'shell', 'kotlin', 'groovy', 'scala',
    'clojure', 'haskell', 'elm', 'erlang', 'ocaml', 'r', 'dart', 'julia', 'elixir', 'crystal',
    'nim', 'reason', 'html', 'css', 'scss', 'less', 'stylus', 'pug', 'handlebars',
    'ejs', 'twig', 'bash', 'sh', 'shell', 'awk', 'vim', 'makefile', 'cmake',]);


const CREDENTIALS_PATH = 'C:/Work/Credentials/credentials.json';
const DATA_PATH = 'C:/Work/Credentials/post-ids.json';

async function downloadFile(file, type) {
    const lastSlash = file.lastIndexOf('/');
    const url = file.substring(0, lastSlash) + "/export?format=" + type;
    const response = await fetch(url);
    let filename = null;
    if (!response.ok) {
        let errorBody = '';
        try {
            // Try to get the response body as text
            errorBody = await response.text();
        } catch (bodyError) {
            console.error('Failed to read error body:', bodyError);
        }
        throw new Error(`HTTP error! status: ${response.status}, body: ${errorBody}`);
    }
    var contentDisposition = response.headers.get('Content-Disposition');
    const utf8Part = contentDisposition.match(/filename\*=UTF-8''(.+)/);
    if (utf8Part && utf8Part[1]) {
        filename = decodeURIComponent(utf8Part[1]);
        filename = filename.substring(0, filename.lastIndexOf('.'));
    }
    return [response, filename];
}

function matchesSelector(selector, element) {
    if (selector.startsWith('#')) {
        return selector.slice(1) === element.id;
    }
    if (selector.startsWith('.')) {
        element.classes = element.classes || (element.attrs.class || '').split(' ');
        return element.classes.includes(selector.slice(1));
    }
    return selector === element.name;
}

function findMatchingRules(element, rules) {
    return rules.filter(rule =>
        (rule.selectors || []).some(selector =>
            selector.split(' ').every(s => matchesSelector(s, element))
        )
    );
}


async function processFile(file) {

    const [zipDownload, _] = await downloadFile(file, 'zip');
    const [textDownload, fileName] = await downloadFile(file, 'txt');
    const text = await textDownload.text();


    let blocks = [];
    for (const match of text.matchAll(/\uEC03(.*?)\uEC02/gs)) {
        const code = match[1].trim();
        let lang = flourite(code, { shiki: true, noUnknown: true }).language;
        if (Prism.languages.hasOwnProperty(lang) == false) {
            lang = "bash";
        }
        const formattedCode = Prism.highlight(code, Prism.languages[lang] || Prism.languages['clike'], lang);

        blocks.push("<hr/><pre class='line-numbers language-" + lang + "'>" +
            "<code class='line-numbers language-" + lang + "'>" +
            formattedCode + "</code></pre><hr/>");
    }

    const zip = new admZip(Buffer.from(await zipDownload.arrayBuffer()));
    const entries = zip.getEntries();
    var htmlText = entries.find(e => path.extname(e.entryName) == '.html')
        .getData().toString('utf8');

    const htmlDoc = new jssoup(htmlText)
    const style = htmlDoc.find('style');

    let styleRules = css.parse(style.text).stylesheet.rules;

    const body = htmlDoc.find('body');

    let tags = [];
    let postId = null;
    let codeSegmentIndex = 0;
    let inCodeSegment = false;
    body.findAll().forEach(e => {
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
    body.findAll('span').forEach(span => {
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

    body.findAll('a').forEach(a => {
        if (a.attrs.hasOwnProperty('href')) {
            // Google Docs put a redirect like: https://www.google.com/url?q=ACTUAL_URL
            var link = new URL(a.attrs.href);
            var q = link.searchParams.get('q');
            if (q) {
                a.attrs.href = q;
            }
        }
    });
    body.findAll('img').forEach(img => {
        if (img.attrs.hasOwnProperty('src')) {
            let src = img.attrs.src;
            let dimensions = extractDimensions(img.attrs.style);
            let imgName = src.split('/').pop();
            let imgData = entries.find(e => e.entryName === 'images/' + imgName).getData();
            let imgType = imgName.split('.').pop();
            let imgSrc = 'data:image/' + imgType + ';base64,' + imgData.toString('base64');
            if (!dimensions || dimensions.width < 200 && dimensions.height < 200) {
                img.replaceWith('<img src="' + imgSrc + '" style="float: right"/>');
            }
            else {
                img.replaceWith('<img src="' + imgSrc + '"/>');
            }
        }
    })

    var currentElement = body.nextElement;
    let cleanHTML = cleanHtml(currentElement, styleRules);
    while (currentElement.nextSibling !== undefined) {
        currentElement = currentElement.nextSibling;
        cleanHTML += cleanHtml(currentElement, styleRules);
    }

    cleanHTML += `\r\n<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/prism/9000.0.1/themes/prism.min.css" integrity="sha512-/mZ1FHPkg6EKcxo0fKXF51ak6Cr2ocgDi5ytaTBjsQZIH/RNs6GF6+oId/vPe3eJB836T36nXwVh/WBl/cWT4w==" crossorigin="anonymous" referrerpolicy="no-referrer" />`;

    if (!postId) {
        postId = readPostId(file);
    }

    return [cleanHTML, postId, tags, fileName];
}

function extractDimensions(cssString) {
    const matches = cssString.match(/width: (\d+\.\d+)px; height: (\d+\.\d+)px;/);
    if (matches) {
        const width = parseFloat(matches[1]);
        const height = parseFloat(matches[2]);
        return { width, height };
    } else {
        return null; // Invalid CSS string format
    }
}

function cleanHtml(htmlElement, styleRules) {
    if (htmlElement.name == "hr") {
        return "<hr/>";
    }
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
    let newStyles = computeNewStyles(htmlElement, modifiers, styleRules);
    modifiers['styles'] = newStyles;
    let href;
    let classes;
    if (htmlElement.name == "table") {
        classes = ' class="table-bordered table-striped" ';
    }
    else {
        classes = '';
    }
    if (htmlElement.name === "a") {
        href = ' href="' + htmlElement.attrs.href + '"';
    } else {
        href = '';
    }
    let newElemStart;
    let newElemEnd;
    if (newStyles) {
        newElemStart = '<' + htmlElement.name + ' style="' + modifiers['styles'] + '"' + href + classes + '>';
        newElemEnd = '</' + htmlElement.name + '>';
    } else {
        newElemStart = '<' + htmlElement.name + href + '>';
        newElemEnd = '</' + htmlElement.name + '>';
    }
    let contents = htmlElement.contents;
    let newContents = '';
    for (let i = 0; i < contents.length; i++) {
        newContents += cleanHtml(contents[i], styleRules);
    }
    if (newContents === "") {
        return '';
    }
    if (htmlElement.name === "span" && !modifiers['styles'] || htmlElement.name === "b") {
        newElemStart = '';
        newElemEnd = '';
    }
    let newHtmlElement = `${modifiers['quote'] ? '<blockquote>' : ''}${newElemStart}${modifiers['bold'] ? '<strong>' : ''}${modifiers['italics'] ? '<em>' : ''}${newContents}${modifiers['italics'] ? '</em>' : ''}${modifiers['bold'] ? '</strong>' : ''}${newElemEnd}${modifiers['quote'] ? '</blockquote>' : ''}`;
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

function applyStyle(styleType, styleValue, htmlElement, modifiers, validStyles) {
    let newStyles = '';

    if (styleType == 'margin-left' && htmlElement.name === "p") {
        modifiers['quote'] = styleValue === '36pt';
    }

    if (validStyles.has(styleType) === false || styleValue === undefined)
        return newStyles;

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

    return newStyles;
}

function computeNewStyles(htmlElement, modifiers, styleRules) {
    const validStyles = computeValidStyles(htmlElement);

    let newStyles = '';
    if (htmlElement.name === "table") {
        newStyles += "width:100%;"
    }
    if (htmlElement.name === "tr") {
        modifiers['bold'] = htmlElement.previousElement.name === "table";
    }
    if (htmlElement.attrs.style) {
        for (let style of htmlElement.attrs.style.split(';')) {
            let parts = style.split(':');
            if (parts.length !== 2)
                continue;
            let styleType = parts[0].trim(), styleValue = parts[1].trim();
            applyStyle(styleType, styleValue, htmlElement, modifiers, validStyles);
        }
    }
    for (let match of findMatchingRules(htmlElement, styleRules)) {
        for (let declaration of match.declarations) {
            if (declaration.type === 'declaration') {
                newStyles += applyStyle(declaration.property, declaration.value, htmlElement, modifiers, validStyles,);
            }
        }
    }

    return newStyles;
}

async function getBlogClient() {
    const content = fs.readFileSync(CREDENTIALS_PATH);
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

function readPostId(key) {
    let data;
    try {
        data = JSON.parse(fs.readFileSync(DATA_PATH, 'utf8'));
    } catch (err) {
        data = {};
    }
    return data[key];
}

function writePostId(key, value) {
    let data;
    try {
        data = JSON.parse(fs.readFileSync(DATA_PATH, 'utf8'));
    } catch (err) {
        data = {};
    }
    data[key] = value;
    fs.writeFileSync(DATA_PATH, JSON.stringify(data, null, 2)); // Pretty print with 2 spaces
}

(async () => {

    const file = process.argv[2];

    const [html, postId, tags, title] = await processFile(file);

    var post = {
        description: html,
        title: title,
        categories: tags,
    };

    const blogClient = await getBlogClient();

    if (postId) {
        await blogClient.editPost(postId, post);
    }
    else {
        const newPostId = await blogClient.createPost(post);
        writePostId(file, newPostId);
        console.log('Created: ', newPostId);
    }
    console.log('Published: ', title);
    exec("start http://ayende.com/blog/")
})()