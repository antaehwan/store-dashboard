// 대시보드의 [지금 갱신] 버튼이 호출하는 엔드포인트.
// GitHub Actions 의 gift-update 워크플로를 실행시킨다.
// 토큰은 Vercel 환경변수 GH_TOKEN 에만 두고 브라우저로는 절대 내보내지 않는다.

const OWNER = "antaehwan";
const REPO = "store-dashboard";
const WORKFLOW = "gift-update.yml";
const COOLDOWN_MS = 45 * 1000;   // 직전 실행 후 최소 간격

function gh(path, token, init = {}) {
  return fetch(`https://api.github.com/repos/${OWNER}/${REPO}${path}`, {
    ...init,
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/vnd.github+json",
      "X-GitHub-Api-Version": "2022-11-28",
      "User-Agent": "gift-dashboard",
      ...(init.headers || {}),
    },
  });
}

module.exports = async (req, res) => {
  res.setHeader("Cache-Control", "no-store");

  if (req.method !== "POST") {
    return res.status(405).json({ ok: false, error: "POST 만 허용됩니다" });
  }

  const token = process.env.GH_TOKEN;
  if (!token) {
    return res.status(500).json({
      ok: false,
      code: "no_token",
      error: "서버에 GH_TOKEN 이 설정되지 않았습니다",
    });
  }

  try {
    // 1) 이미 돌고 있거나 방금 끝났으면 중복 실행하지 않는다
    const listRes = await gh(
      `/actions/workflows/${WORKFLOW}/runs?per_page=5`, token);
    if (listRes.ok) {
      const { workflow_runs: runs = [] } = await listRes.json();
      const busy = runs.find(
        (r) => r.status === "in_progress" || r.status === "queued");
      if (busy) {
        return res.status(200).json({ ok: true, status: "running" });
      }
      const last = runs[0];
      if (last && Date.now() - new Date(last.updated_at).getTime() < COOLDOWN_MS) {
        return res.status(200).json({ ok: true, status: "cooldown" });
      }
    }

    // 2) 워크플로 실행
    const runRes = await gh(
      `/actions/workflows/${WORKFLOW}/dispatches`, token,
      { method: "POST", body: JSON.stringify({ ref: "main" }) });

    if (runRes.status === 204) {
      return res.status(200).json({ ok: true, status: "dispatched" });
    }
    const body = await runRes.text();
    return res.status(502).json({
      ok: false,
      code: "dispatch_failed",
      error: `GitHub 응답 ${runRes.status}`,
      detail: body.slice(0, 300),
    });
  } catch (e) {
    return res.status(500).json({ ok: false, error: String(e).slice(0, 300) });
  }
};
