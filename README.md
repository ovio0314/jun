# 오비오 인사노무 브리핑센터

이 저장소는 Sites 프로젝트 `ovio-hr-labor-briefing`의 원본 소스입니다.
향후 수정 전에는 `PROJECT_MEMORY.md`를 먼저 확인하고 기존 데이터와 공개 설정을 유지합니다.

## 편집 기준

- Sites project ID: `appgprj_6a993c24256481919510c0a99761a2cd`
- 공개 URL: `https://ovio-hr-labor-briefing.ovio-4448.chatgpt.site`
- 편집 원본: `worker/index.js`
- 생성 결과물: `dist/index.html`
- 기존 법령·판례·뉴스·AI·지원사업 데이터는 삭제하거나 초기화하지 않습니다.

## 기술 메모

Use this starter for a static microsite, click counter, or simple internal UI whose state is browser-scoped. It has no dependencies and needs no install.

Edit `worker/index.js`. Use the Sites checkpoint when a coherent milestone is ready to inspect or share; the remote builder then runs the checked-in build and validation scripts. Do not run them as a normal pre-checkpoint step.

The build copies only `worker/index.js` and `.openai/hosting.json`. Do not add standalone asset files. Embed any essential raster bytes in `worker/index.js` and serve or reference them as a data URL.

For targeted diagnosis after a remote build failure, the same commands are available in the Sites Linux environment:

```sh
bash scripts/build.sh
node scripts/validate-artifact.mjs
```

The deterministic build produces:

```text
dist/
├── .openai/
│   └── hosting.json
└── server/
    └── index.js
```

`dist/server/index.js` is an ES module with a default export containing `fetch(request, env, ctx)`. Edit `worker/index.js`, not the generated file under `dist/`.
