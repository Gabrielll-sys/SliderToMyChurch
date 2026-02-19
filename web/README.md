# Slides To My Church Web

MVP em Next.js para:

- editar seções da celebração;
- importar leituras por data (`/api/liturgia`);
- gerar e baixar `.pptx` (`/api/generate`).

## Rodar local

```bash
cd web
npm install
npm run dev
```

Abra `http://localhost:3000`.

## Comandos úteis

```bash
npm run lint
npm run build
```

## Deploy no Vercel

1. Suba o repositório para o GitHub.
2. No Vercel, clique em `New Project` e selecione este repositório.
3. Em `Root Directory`, selecione `web`.
4. Framework: `Next.js`.
5. Clique em `Deploy`.

## Estrutura principal

- `src/app/page.tsx`: interface web de edição.
- `src/app/api/liturgia/route.ts`: proxy/normalização da API litúrgica externa.
- `src/app/api/generate/route.ts`: geração de PowerPoint com `pptxgenjs`.
- `src/lib/missa.ts`: tipos, defaults e utilitários compartilhados.
