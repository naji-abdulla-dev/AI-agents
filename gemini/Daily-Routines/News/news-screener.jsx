import { useState, useEffect, useCallback } from "react";

const TOPICS = [
  { id: "tech", label: "Technology", emoji: "💻", color: "#06b6d4" },
  { id: "ai", label: "AI", emoji: "🤖", color: "#8b5cf6" },
  { id: "programming", label: "Programming", emoji: "⌨️", color: "#10b981" },
  { id: "finance", label: "Finance", emoji: "📈", color: "#f59e0b" },
  { id: "economy", label: "Economy", emoji: "🌐", color: "#ef4444" },
];

const LOADING_QUIPS = [
  "Shaking down the internet for hot takes...",
  "Bribing search engines for the good stuff...",
  "Sifting signal from noise (mostly noise)...",
  "Claude is reading the news so you don't have to...",
  "Asking the algorithm nicely for 20 bangers...",
  "Hunting for stories that won't put you to sleep...",
  "Curating your daily dose of \"oh interesting\"...",
  "Separating clickbait from click-worthy...",
];

const SYSTEM_PROMPT = `You are a news screener. Your job is to find the 20 most important/interesting news items from today or the past 24-48 hours across these topics: Technology, AI, Programming, Finance, Economy.

Return ONLY a JSON array of exactly 20 news objects. No markdown, no backticks, no preamble. Just raw JSON.

Each object must have:
- "title": concise headline (your own wording, max 15 words)
- "summary": 1-2 sentence summary in your own words
- "topic": one of "tech", "ai", "programming", "finance", "economy"
- "source": the publication/site name
- "url": the actual URL if found
- "spiciness": 1-5 rating of how interesting/important this is (5 = massive deal)

Distribution: aim for roughly 4 items per topic, but if one topic has huge news, flex.
Sort by spiciness descending.
Prioritize: recency, impact, relevance to practitioners (not just casual readers).
Avoid paywalled-only content where possible.`;

function NewsCard({ item, index }) {
  const topic = TOPICS.find((t) => t.id === item.topic) || TOPICS[0];
  const spice = "🔥".repeat(item.spiciness || 1);

  return (
    <div
      style={{
        background: "var(--card-bg)",
        borderRadius: "12px",
        padding: "20px",
        border: "1px solid var(--border)",
        transition: "all 0.25s ease",
        cursor: "pointer",
        position: "relative",
        overflow: "hidden",
        animationDelay: `${index * 60}ms`,
      }}
      className="news-card"
      onClick={() => item.url && window.open(item.url, "_blank")}
    >
      <div
        style={{
          position: "absolute",
          top: 0,
          left: 0,
          right: 0,
          height: "3px",
          background: topic.color,
        }}
      />
      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "flex-start",
          marginBottom: "10px",
          gap: "12px",
        }}
      >
        <span
          style={{
            fontSize: "11px",
            fontWeight: 700,
            textTransform: "uppercase",
            letterSpacing: "0.08em",
            color: topic.color,
            background: `${topic.color}18`,
            padding: "3px 10px",
            borderRadius: "100px",
            whiteSpace: "nowrap",
            flexShrink: 0,
          }}
        >
          {topic.emoji} {topic.label}
        </span>
        <span style={{ fontSize: "14px", flexShrink: 0 }} title="Spiciness rating">
          {spice}
        </span>
      </div>
      <h3
        style={{
          fontSize: "16px",
          fontWeight: 700,
          lineHeight: 1.35,
          margin: "0 0 8px 0",
          color: "var(--text-primary)",
        }}
      >
        {item.title}
      </h3>
      <p
        style={{
          fontSize: "13.5px",
          lineHeight: 1.55,
          color: "var(--text-secondary)",
          margin: "0 0 12px 0",
        }}
      >
        {item.summary}
      </p>
      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
        }}
      >
        <span
          style={{
            fontSize: "12px",
            color: "var(--text-muted)",
            fontStyle: "italic",
          }}
        >
          {item.source || "—"}
        </span>
        {item.url && (
          <span
            style={{
              fontSize: "12px",
              color: topic.color,
              fontWeight: 600,
            }}
          >
            Read →
          </span>
        )}
      </div>
    </div>
  );
}

function TopicFilter({ activeTopics, onToggle }) {
  return (
    <div style={{ display: "flex", gap: "8px", flexWrap: "wrap" }}>
      {TOPICS.map((t) => {
        const active = activeTopics.includes(t.id);
        return (
          <button
            key={t.id}
            onClick={() => onToggle(t.id)}
            style={{
              padding: "6px 14px",
              borderRadius: "100px",
              border: `1.5px solid ${active ? t.color : "var(--border)"}`,
              background: active ? `${t.color}20` : "transparent",
              color: active ? t.color : "var(--text-muted)",
              fontSize: "13px",
              fontWeight: 600,
              cursor: "pointer",
              transition: "all 0.2s ease",
              fontFamily: "inherit",
            }}
          >
            {t.emoji} {t.label}
          </button>
        );
      })}
    </div>
  );
}

function StatsBar({ items }) {
  const counts = {};
  TOPICS.forEach((t) => (counts[t.id] = 0));
  items.forEach((item) => {
    if (counts[item.topic] !== undefined) counts[item.topic]++;
  });
  const avgSpice = items.length
    ? (items.reduce((s, i) => s + (i.spiciness || 1), 0) / items.length).toFixed(1)
    : 0;

  return (
    <div
      style={{
        display: "flex",
        gap: "16px",
        flexWrap: "wrap",
        padding: "14px 20px",
        background: "var(--card-bg)",
        borderRadius: "12px",
        border: "1px solid var(--border)",
        alignItems: "center",
        fontSize: "13px",
      }}
    >
      {TOPICS.map((t) => (
        <span key={t.id} style={{ color: t.color, fontWeight: 600 }}>
          {t.emoji} {counts[t.id]}
        </span>
      ))}
      <span style={{ color: "var(--text-muted)", marginLeft: "auto" }}>
        Avg spiciness: {avgSpice} 🔥
      </span>
    </div>
  );
}

export default function NewsScreener() {
  const [items, setItems] = useState([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [quipIdx, setQuipIdx] = useState(0);
  const [activeTopics, setActiveTopics] = useState(TOPICS.map((t) => t.id));
  const [lastFetched, setLastFetched] = useState(null);

  useEffect(() => {
    if (!loading) return;
    const iv = setInterval(() => {
      setQuipIdx((i) => (i + 1) % LOADING_QUIPS.length);
    }, 3200);
    return () => clearInterval(iv);
  }, [loading]);

  const fetchNews = useCallback(async () => {
    setLoading(true);
    setError(null);
    setItems([]);

    try {
      const today = new Date().toLocaleDateString("en-US", {
        weekday: "long",
        year: "numeric",
        month: "long",
        day: "numeric",
      });

      const response = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "x-api-key": "YOUR_API_KEY_HERE",
          "anthropic-version": "2023-06-01",
          "anthropic-dangerously-allow-browser": "true"
        },
        body: JSON.stringify({
          model: "claude-sonnet-4-20250514",
          max_tokens: 4000,
          system: SYSTEM_PROMPT,
          tools: [{ type: "web_search_20250305", name: "web_search" }],
          messages: [
            {
              role: "user",
              content: `Today is ${today}. Search for the latest news across Technology, AI, Programming, Finance, and Economy. Find the 20 most important stories from the past 24-48 hours. Return ONLY a raw JSON array.`,
            },
          ],
        }),
      });

      if (!response.ok) {
        const errBody = await response.text();
        throw new Error(`API error ${response.status}: ${errBody.slice(0, 200)}`);
      }

      const data = await response.json();

      // Extract text content from response
      const textBlocks = data.content
        .filter((b) => b.type === "text")
        .map((b) => b.text)
        .join("\n");

      // Try to parse JSON from the response
      let parsed = null;
      // Try raw parse first
      try {
        parsed = JSON.parse(textBlocks.trim());
      } catch {
        // Try to extract JSON array from the text
        const match = textBlocks.match(/\[[\s\S]*\]/);
        if (match) {
          parsed = JSON.parse(match[0]);
        }
      }

      if (!parsed || !Array.isArray(parsed)) {
        throw new Error("Claude didn't return a valid JSON array. Response: " + textBlocks.slice(0, 300));
      }

      setItems(parsed);
      setLastFetched(new Date());
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  }, []);

  const toggleTopic = (id) => {
    setActiveTopics((prev) =>
      prev.includes(id) ? prev.filter((t) => t !== id) : [...prev, id]
    );
  };

  const filtered = items.filter((i) => activeTopics.includes(i.topic));

  return (
    <div
      style={{
        fontFamily: "'IBM Plex Sans', 'SF Pro Display', system-ui, sans-serif",
        minHeight: "100vh",
        background: "var(--bg)",
        color: "var(--text-primary)",
        padding: "32px 24px",
        maxWidth: "900px",
        margin: "0 auto",
      }}
    >
      <style>{`
        :root {
          --bg: #0c0f14;
          --card-bg: #151922;
          --border: #1f2937;
          --text-primary: #e8ecf1;
          --text-secondary: #94a3b8;
          --text-muted: #64748b;
          --accent: #06b6d4;
        }
        @media (prefers-color-scheme: light) {
          :root {
            --bg: #f8fafc;
            --card-bg: #ffffff;
            --border: #e2e8f0;
            --text-primary: #0f172a;
            --text-secondary: #475569;
            --text-muted: #94a3b8;
            --accent: #0891b2;
          }
        }
        @keyframes fadeUp {
          from { opacity: 0; transform: translateY(16px); }
          to { opacity: 1; transform: translateY(0); }
        }
        .news-card {
          animation: fadeUp 0.4s ease both;
        }
        .news-card:hover {
          border-color: var(--accent) !important;
          transform: translateY(-2px);
          box-shadow: 0 8px 24px rgba(0,0,0,0.15);
        }
        @keyframes pulse {
          0%, 100% { opacity: 1; }
          50% { opacity: 0.5; }
        }
        .loading-dot {
          animation: pulse 1.4s ease-in-out infinite;
        }
        @keyframes spin {
          to { transform: rotate(360deg); }
        }
      `}</style>

      {/* Header */}
      <div style={{ marginBottom: "28px" }}>
        <div style={{ display: "flex", alignItems: "center", gap: "12px", marginBottom: "6px" }}>
          <span style={{ fontSize: "28px" }}>📡</span>
          <h1
            style={{
              fontSize: "26px",
              fontWeight: 800,
              margin: 0,
              background: "linear-gradient(135deg, #06b6d4, #8b5cf6)",
              WebkitBackgroundClip: "text",
              WebkitTextFillColor: "transparent",
              letterSpacing: "-0.02em",
            }}
          >
            Daily News Screener
          </h1>
        </div>
        <p style={{ color: "var(--text-muted)", fontSize: "14px", margin: "4px 0 0 40px" }}>
          20 curated stories across Tech · AI · Programming · Finance · Economy
        </p>
      </div>

      {/* Controls */}
      <div
        style={{
          display: "flex",
          flexWrap: "wrap",
          gap: "12px",
          alignItems: "center",
          marginBottom: "20px",
        }}
      >
        <button
          onClick={fetchNews}
          disabled={loading}
          style={{
            padding: "10px 24px",
            borderRadius: "10px",
            border: "none",
            background: loading ? "var(--border)" : "linear-gradient(135deg, #06b6d4, #8b5cf6)",
            color: "#fff",
            fontSize: "14px",
            fontWeight: 700,
            cursor: loading ? "not-allowed" : "pointer",
            fontFamily: "inherit",
            display: "flex",
            alignItems: "center",
            gap: "8px",
            transition: "all 0.2s ease",
          }}
        >
          {loading ? (
            <>
              <span
                style={{
                  width: "16px",
                  height: "16px",
                  border: "2px solid rgba(255,255,255,0.3)",
                  borderTopColor: "#fff",
                  borderRadius: "50%",
                  display: "inline-block",
                  animation: "spin 0.8s linear infinite",
                }}
              />
              Scanning...
            </>
          ) : (
            <>⚡ Fetch Today's News</>
          )}
        </button>
        {lastFetched && (
          <span style={{ fontSize: "12px", color: "var(--text-muted)" }}>
            Last fetched:{" "}
            {lastFetched.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" })}
          </span>
        )}
      </div>

      {/* Loading state */}
      {loading && (
        <div
          style={{
            textAlign: "center",
            padding: "60px 20px",
            color: "var(--text-secondary)",
          }}
        >
          <div style={{ fontSize: "48px", marginBottom: "16px" }}>🔍</div>
          <p style={{ fontSize: "16px", fontWeight: 600, marginBottom: "8px" }}>
            {LOADING_QUIPS[quipIdx]}
          </p>
          <div style={{ display: "flex", justifyContent: "center", gap: "6px", marginTop: "16px" }}>
            {[0, 1, 2].map((i) => (
              <span
                key={i}
                className="loading-dot"
                style={{
                  width: "8px",
                  height: "8px",
                  borderRadius: "50%",
                  background: "var(--accent)",
                  animationDelay: `${i * 0.2}s`,
                }}
              />
            ))}
          </div>
        </div>
      )}

      {/* Error */}
      {error && (
        <div
          style={{
            padding: "16px 20px",
            borderRadius: "12px",
            background: "#ef44441a",
            border: "1px solid #ef444440",
            color: "#f87171",
            fontSize: "14px",
            marginBottom: "20px",
            lineHeight: 1.5,
          }}
        >
          <strong>Oops!</strong> {error}
        </div>
      )}

      {/* Results */}
      {items.length > 0 && !loading && (
        <>
          <StatsBar items={items} />
          <div style={{ margin: "16px 0" }}>
            <TopicFilter activeTopics={activeTopics} onToggle={toggleTopic} />
          </div>
          <div
            style={{
              display: "grid",
              gridTemplateColumns: "repeat(auto-fill, minmax(380px, 1fr))",
              gap: "14px",
              marginTop: "8px",
            }}
          >
            {filtered.map((item, i) => (
              <NewsCard key={i} item={item} index={i} />
            ))}
          </div>
          {filtered.length === 0 && (
            <p
              style={{
                textAlign: "center",
                color: "var(--text-muted)",
                padding: "40px",
                fontSize: "14px",
              }}
            >
              No stories match your filters. Toggle some topics above!
            </p>
          )}
        </>
      )}

      {/* Empty state */}
      {!loading && items.length === 0 && !error && (
        <div
          style={{
            textAlign: "center",
            padding: "80px 20px",
            color: "var(--text-muted)",
          }}
        >
          <div style={{ fontSize: "56px", marginBottom: "16px" }}>📰</div>
          <p style={{ fontSize: "18px", fontWeight: 600, color: "var(--text-secondary)" }}>
            Your daily briefing awaits
          </p>
          <p style={{ fontSize: "14px", marginTop: "8px" }}>
            Hit the button above and Claude will fetch & curate 20 stories for you.
          </p>
        </div>
      )}
    </div>
  );
}
