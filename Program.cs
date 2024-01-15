using Google.Apis.Auth.OAuth2;
using Google.Apis.Docs.v1;
using Google.Apis.Docs.v1.Data;
using Google.Apis.Drive.v3;
using Google.Apis.Json;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using HtmlAgilityPack;
using System.Diagnostics;
using System.IO.Compression;
using System.Web;

var arg = Environment.GetCommandLineArgs()[1];

var documentId = arg;
if(Uri.TryCreate(arg, UriKind.Absolute, out var uri))
{
    documentId = uri.AbsolutePath.Split('/')[^2]; // the one after the /edit
}

var state = NewtonsoftJsonSerializer.Instance.Deserialize<Configuration>(File.ReadAllText(@"C:\Work\Credentials\blog.json"));


var publisher = new PublishingPlatform(state);

publisher.Publish(documentId);

Process.Start("cmd", "/c start https://ayende.com/blog/admin");

public class Configuration
{
    public ClientSecrets Secrets;
    public BlogCredentials Blog;
    public class BlogCredentials
    {
        public string Username;
        public string Password;
    }
}


public class PublishingPlatform
{
    private readonly DocsService GoogleDocs;
    private readonly DriveService GoogleDrive;
    private readonly MetaWeblogClient.Client _blogClient;

    public PublishingPlatform(Configuration cfg)
    {
        var blogInfo = new  MetaWeblogClient.BlogConnectionInfo(
         "https://ayende.com/blog",
         "https://ayende.com/blog/Services/MetaWeblogAPI.ashx",
         "ayende.com", cfg.Blog.Username, cfg.Blog.Password);
        _blogClient = new MetaWeblogClient.Client(blogInfo);

        var initializer = new BaseClientService.Initializer
        {
            HttpClientInitializer = GoogleWebAuthorizationBroker.AuthorizeAsync(
              cfg.Secrets,
              new[] { DocsService.Scope.Documents, DriveService.Scope.DriveReadonly },
              "user", CancellationToken.None,
              new FileDataStore("blog.ayende.com")
          ).Result
        };

        GoogleDocs = new DocsService(initializer);
        GoogleDrive = new DriveService(initializer);
    }

    public void Publish(string documentId)
    {
        using var file = GoogleDrive.Files.Export(documentId, "application/zip").ExecuteAsStream();
        var zip = new ZipArchive(file, ZipArchiveMode.Read);

        var doc = GoogleDocs.Documents.Get(documentId).Execute();
        var title = doc.Title;

        var htmlFile = zip.Entries.First(e => Path.GetExtension(e.Name).ToLower() == ".html");
        using var stream = htmlFile.Open();
        var htmlDoc = new HtmlDocument();
        htmlDoc.Load(stream);
        var body = htmlDoc.DocumentNode.SelectSingleNode("//body");

        var (postId, tags) = ReadPostIdAndTags(body);

        UpdateLinks(body);
        StripCodeHeader(body);
        UploadImages(zip, body, GenerateSlug(title));

        string post = GetPostContents(htmlDoc, body);

        if (postId != null)
        {
            _blogClient.EditPost(postId, title, post, tags, true);
            return;
        }

        postId = _blogClient.NewPost(title, post, tags, true, null);

        var update = new BatchUpdateDocumentRequest
        {
            Requests = [new Request
            {
                InsertText = new InsertTextRequest
                {
                    Text = $"PostId: {postId}\r\n",
                    Location = new Location
                    {
                        Index = 1,
                    }
                },
            }]
        };

        GoogleDocs.Documents.BatchUpdate(update, documentId).Execute();
    }

    private void StripCodeHeader(HtmlNode body)
    {
        foreach (var remove in body.SelectNodes("//span[text()='&#60419;']").ToArray())
        {
            remove.Remove();
        }
        foreach (var remove in body.SelectNodes("//span[text()='&#60418;']").ToArray())
        {
            remove.Remove();
        }
    }

    private static string GetPostContents(HtmlDocument htmlDoc, HtmlNode body)
    {
        // we use the @scope element to ensure that the document style doesn't "leak" outside
        var style = htmlDoc.DocumentNode.SelectSingleNode("//head/style[@type='text/css']").InnerText;
        var post = "<style>@scope {" + style + "}</style> " + body.InnerHtml;
        return post;
    }

    private static void UpdateLinks(HtmlNode body)
    {
        // Google Docs put a redirect like: https://www.google.com/url?q=ACTUAL_URL
        foreach (var link in body.SelectNodes("//a[@href]").ToArray())
        {
            var href = new Uri(link.Attributes["href"].Value);
            var url = HttpUtility.ParseQueryString(href.Query)["q"];
            if (url != null)
            {
                link.Attributes["href"].Value = url;
            }
        }
    }

    private static (string? postId, List<string> tags) ReadPostIdAndTags(HtmlNode body)
    {
        string? postId = null;
        var tags = new List<string>();
        foreach (var span in body.SelectNodes("//span"))
        {
            var text = span.InnerText.Trim();
            const string TagsPrefix = "Tags:";
            const string PostIdPrefix = "PostId:";
            if (text.StartsWith(TagsPrefix, StringComparison.OrdinalIgnoreCase))
            {
                tags.AddRange(text.Substring(TagsPrefix.Length).Split(","));
                RemoveElement(span);
            }
            else if (text.StartsWith(PostIdPrefix, StringComparison.OrdinalIgnoreCase))
            {
                postId = text.Substring(PostIdPrefix.Length).Trim();
                RemoveElement(span);
            }
        }
        // after we removed post id & tags, trim the empty lines
        while (body.FirstChild.InnerText.Trim() is "&nbsp;" or "")
        {
            body.RemoveChild(body.FirstChild);
        }
        return (postId, tags);
    }

    private static void RemoveElement(HtmlNode element)
    {
        do
        {
            var parent = element.ParentNode;
            parent.RemoveChild(element);
            element = parent;
        } while (element?.ChildNodes?.Count == 0);
    }

    private void UploadImages(ZipArchive zip, HtmlNode body, string slug)
    {
        var mapping = new Dictionary<string, string>();
        foreach (var image in zip.Entries.Where(x => Path.GetDirectoryName(x.FullName) == "images"))
        {
            var type = Path.GetExtension(image.Name).ToLower() switch
            {
                ".png" => "image/png",
                ".jpg" or "jpeg" => "image/jpg",
                _ => "application/octet-stream"
            };
            using var contents = image.Open();
            var ms = new MemoryStream();
            contents.CopyTo(ms);
            var bytes = ms.ToArray();
            var result = _blogClient.NewMediaObject(slug + "/" + Path.GetFileName(image.Name), type, bytes);
            mapping[image.FullName] = new UriBuilder { Path = result.URL }.Uri.AbsolutePath;
        }
        foreach (var img in body.SelectNodes("//img[@src]").ToArray())
        {
            if (mapping.TryGetValue(img.Attributes["src"].Value, out var path))
            {
                img.Attributes["src"].Value = path;
            }
        }
    }

    private static string GenerateSlug(string title)
    {
        var slug = title.Replace(" ", "");
        foreach (var ch in Path.GetInvalidFileNameChars())
        {
            slug = slug.Replace(ch, '-');
        }

        return slug;
    }
}
