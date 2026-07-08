# html2pptx-service

A lightweight Python microservice that converts structured HTML slides into fully formatted `.pptx` PowerPoint files. Designed to be deployed as a Docker container and called via HTTP POST.

---

## How It Works

1. You POST an HTML string to the service
2. The service parses the HTML, detects slide layouts from CSS classes and `data-*` attributes
3. It generates a styled `.pptx` file using `python-pptx`
4. The `.pptx` binary is returned directly in the HTTP response

The service is stateless — every request is independent.

---

## Supported Slide Layouts

Every slide must be a `<section class="slide">` element. The service detects the layout from attributes.

### Layout 1 — Title Slide
First slide of the presentation. Detected automatically when a `<section class="slide">` has an `<h1>` and a `<p>` but no `<ul>`.

```html
<section class="slide">
  <h1>Presentation Title</h1>
  <p>One-liner subtitle or tagline.</p>
</section>
```

### Layout 2 — Bullets Slide
Default layout. Used when no `data-chart` or `data-layout` attribute is present.

```html
<section class="slide">
  <h1>Slide Title</h1>
  <ul>
    <li>Point one</li>
    <li>Point two</li>
    <li>Point three</li>
  </ul>
</section>
```

### Layout 3 — Bar Chart Slide
Renders a clustered column chart. Automatically places a data table below the chart (≤6 categories) or beside it (>6 categories).

```html
<section class="slide"
  data-chart="bar"
  data-chart-title="Chart Title"
  data-categories="Category A,Category B,Category C"
  data-series-name="Series Label"
  data-values="45,30,25">
  <h1>Slide Title</h1>
  <p>One sentence insight about the data.</p>
</section>
```

### Layout 4 — Pie Chart Slide
Renders a pie chart with percentage labels. Data table always appears to the right of the chart.

```html
<section class="slide"
  data-chart="pie"
  data-chart-title="Chart Title"
  data-categories="Segment A,Segment B,Segment C"
  data-values="50,30,20">
  <h1>Slide Title</h1>
  <p>One sentence insight about the data.</p>
</section>
```

### Layout 5 — Two Column Slide
Side-by-side comparison layout with two styled cards.

```html
<section class="slide" data-layout="two-column">
  <h1>Slide Title</h1>
  <div class="col-left">
    <h2>Left Heading</h2>
    <ul>
      <li>Point one</li>
      <li>Point two</li>
    </ul>
  </div>
  <div class="col-right">
    <h2>Right Heading</h2>
    <ul>
      <li>Point one</li>
      <li>Point two</li>
    </ul>
  </div>
</section>
```

### Layout 6 — Mixed Slide (Bullets + Chart + Table)
Bullets on the left card, chart on the top-right, data table on the bottom-right. Supports both `bar` and `pie` chart types.

```html
<section class="slide"
  data-layout="mixed"
  data-chart="bar"
  data-chart-title="Chart Title"
  data-categories="Category A,Category B,Category C"
  data-series-name="Series Label"
  data-values="45,30,25">
  <h1>Slide Title</h1>
  <div class="col-left">
    <ul>
      <li>Key insight one</li>
      <li>Key insight two</li>
      <li>Key insight three</li>
    </ul>
  </div>
</section>
```

---

## Data Attribute Reference

| Attribute | Required for | Description |
|---|---|---|
| `data-chart` | Layouts 3, 4, 6 | Chart type: `"bar"` or `"pie"` |
| `data-chart-title` | Layouts 3, 4, 6 | Title shown inside the chart |
| `data-categories` | Layouts 3, 4, 6 | Comma-separated category labels |
| `data-series-name` | Layouts 3, 6 | Series label shown in legend and table header |
| `data-values` | Layouts 3, 4, 6 | Comma-separated numeric values (integers or decimals, no units) |
| `data-layout` | Layouts 5, 6 | Layout override: `"two-column"` or `"mixed"` |

---

## API Reference

### `POST /`

Accepts raw HTML and returns a `.pptx` binary.

**Request**

| Field | Value |
|---|---|
| Method | `POST` |
| Content-Type | `text/html` |
| Body | Raw HTML string containing `<section class="slide">` elements |

The service also accepts base64-encoded HTML — it will attempt to decode it automatically and fall back to raw if decoding fails.

**Response**

| Field | Value |
|---|---|| Status | `200 OK` |
| Content-Type | `application/vnd.openxmlformats-officedocument.presentationml.presentation` |
| Content-Disposition | `attachment; filename="presentation.pptx"` |
| Body | Binary `.pptx` file |

### `OPTIONS /`

CORS preflight endpoint. Returns `200 OK` with appropriate headers. Required for browser-based clients.

---

## Environment Variables

| Variable | Required | Default | Description |
|---|---|---|---|
| `HTML2PPTX_PORT` | No | `8080` | Port the HTTP server listens on |

---

## Deploying with Docker

### Prerequisites
- Docker installed on the target machine
- Port `8080` available (or configure a different port via env var)

### Build and Run Locally

```bash
# Clone the repository
git clone https://github.com/your-org/html2pptx-service.git
cd html2pptx-service

# Build the image
docker build -t html2pptx-service .

# Run the container
docker run -d \
  --name html2pptx \
  -p 8080:8080 \
  -e HTML2PPTX_PORT=8080 \
  html2pptx-service
```

The service will be available at `http://localhost:8080`.

### Using Docker Compose

Create a `docker-compose.yml` alongside your other services:

```yaml
services:
  html2pptx:
    build: ./html2pptx-service   # path to this repo
    ports:
      - "8080:8080"
    environment:
      - HTML2PPTX_PORT=8080
    restart: unless-stopped
```

Then run:

```bash
docker compose up -d html2pptx
```

---

## Deploying on Render.com

1. Push this repository to GitHub
2. Go to [render.com](https://render.com) → **New → Web Service**
3. Connect your GitHub repository
4. Configure the service:

| Setting | Value |
|---|---|
| **Language** | Docker |
| **Branch** | `main` (or `master`) |
| **Region** | Your preferred region |
| **Instance Type** | Free (sufficient for low-moderate traffic) |

5. Under **Environment Variables**, add:

| Key | Value |
|---|---|
| `HTML2PPTX_PORT` | `8080` |

6. Click **Deploy Web Service**

Render will build the Docker image and deploy it. Once live, your service URL will be:
```
https://html2pptx-service.onrender.com
```

> **Note:** Render free-tier instances spin down after 15 minutes of inactivity. The first request after a cold start may take 30–60 seconds. Upgrade to a paid instance if you need always-on availability.

---

## Deploying on a VPS / Linux Server

```bash
# Install Docker if not already installed
curl -fsSL https://get.docker.com | sh

# Clone the repo
git clone https://github.com/your-org/html2pptx-service.git
cd html2pptx-service

# Build and run
docker build -t html2pptx-service .
docker run -d \
  --name html2pptx \
  --restart unless-stopped \
  -p 8080:8080 \
  -e HTML2PPTX_PORT=8080 \
  html2pptx-service

# Verify it's running
docker ps
docker logs html2pptx
```

To expose it publicly, put it behind an Nginx reverse proxy with SSL.

---

## Testing the Service

### Quick health check (expect a 501 — GET is not supported, which confirms the server is up)
There is an attached test.html file in the repo for performing tests.

```bash
curl http://localhost:8080
# Expected: "Error response — Unsupported method ('GET')"
```

### Test with an HTML file (Git Bash / Linux / Mac)

Save your HTML to a file (e.g. `test.html`) then run:

```bash
curl -X POST http://localhost:8080 \
  -H "Content-Type: text/html" \
  --data-binary @test.html \
  --output presentation.pptx
```

Replace `localhost:8080` with your deployed URL if testing a remote instance:

```bash
curl -X POST https://html2pptx-service.onrender.com \
  -H "Content-Type: text/html" \
  --data-binary @test.html \
  --output presentation.pptx
```

### Test with an inline HTML string

```bash
curl -X POST http://localhost:8080 \
  -H "Content-Type: text/html" \
  -d '<html><body>
    <section class="slide"><h1>Test</h1><p>Subtitle here.</p></section>
    <section class="slide"><h1>Bullets</h1><ul><li>Point one</li><li>Point two</li></ul></section>
  </body></html>' \
  --output test.pptx
```

A valid `.pptx` file will be saved. Open it in PowerPoint or LibreOffice to verify.

### Expected output size

A typical 8-slide presentation with 3 charts returns a file of **25–60 KB**.

---

## Calling from JavaScript / TypeScript

```typescript
async function generatePptx(html: string): Promise<void> {
  const response = await fetch('https://html2pptx-service.onrender.com', {
    method: 'POST',
    headers: { 'Content-Type': 'text/html' },
    body: html,
  });

  if (!response.ok) throw new Error(`Service error: ${response.status}`);

  const blob = await response.blob();
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'presentation.pptx';
  a.click();
  URL.revokeObjectURL(url);
}
```

---

## Calling from Python

```python
import requests

with open('slides.html', 'r') as f:
    html = f.read()

response = requests.post(
    'https://html2pptx-service.onrender.com',
    headers={'Content-Type': 'text/html'},
    data=html.encode('utf-8')
)

with open('presentation.pptx', 'wb') as f:
    f.write(response.content)
```

---

## File Structure

```
html2pptx-service/
├── main.py            # HTTP server + all slide builders
├── requirements.txt   # Python dependencies
├── Dockerfile         # Container definition
└── README.md          # This file
```

---

## Dependencies

Defined in `requirements.txt`:

```
python-pptx
beautifulsoup4
lxml
```

All dependencies are installed automatically during the Docker build. No additional setup is required.

---

## Troubleshooting

**502 Bad Gateway on Render**
The service crashed on startup or during a request. Check Render logs under your service → **Logs** tab for the Python traceback.

**`No slides detected` in the output PPTX**
The HTML you sent does not contain any `<section class="slide">` elements. Make sure every slide uses exactly that class.

**Chart shows no data / empty chart**
Check that `data-categories` and `data-values` have the same number of comma-separated items and that all values are plain numbers with no units (e.g. `45,30,25` not `45%,30%,25%`).

**Cold start delay on Render free tier**
The first request after inactivity takes 30–60 seconds. Subsequent requests are fast. Upgrade to a paid Render instance to eliminate cold starts.

**Port already in use**
Change the host port mapping: `-p 9090:8080` will expose the service on port 9090 while the container still listens on 8080 internally.

---

## License

MIT
