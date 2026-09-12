# Input format

The script accepts UTF-8 JSON. `archive_name` and `articles` are required.

```json
{
  "archive_name": "公众号文章归档",
  "articles": [
    {
      "id": "01",
      "label": "活动报道",
      "url": "https://mp.weixin.qq.com/s?..."
    },
    {
      "id": "02",
      "label": "成果展示",
      "url": "https://mp.weixin.qq.com/s/<token>",
      "html_path": "C:/optional/cache/article-02.html"
    }
  ]
}
```

## Fields

- `archive_name` (string): Name of the output folder and ZIP. It must be a safe leaf name, not a path.
- `articles` (array): Ordered article definitions. Duplicate IDs are rejected.
- `id` (string): User-facing identifier. Preserve leading zeroes and gaps.
- `label` (string): Desired folder and DOCX label. Invalid Windows filename characters are replaced with `_`.
- `url` (string): Original public WeChat URL; it is preserved as the source hyperlink.
- `html_path` (optional string): Local HTML snapshot to parse instead of fetching. The snapshot is input only and is never copied into the deliverable.

For a pasted list, convert lines such as `01-标题` followed by a URL into this schema. Never infer or invent a missing URL.

