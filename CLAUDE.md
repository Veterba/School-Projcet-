# CLAUDE.md

## Working agreement
- If you are unsure about anything, ASK before acting.
- Do not implement something you are not confident about.

## Project type
Frontend-only React project (React + Vite). No backend.
Focus: aesthetic design and front-end quality.

## Build & Deploy
- Dev server: `npm run dev`
- Build:      `npm run build`   (outputs to /dist)
- Preview:   `npm run preview`
- Deploy:    `npm run deploy`   (gh-pages -d dist → gh-pages branch)

IMPORTANT:
- vite.config.js MUST set  base: '/<repo-name>/'  or GitHub Pages
  serves a blank page (assets 404).
- If using react-router, use HashRouter (BrowserRouter breaks on
  page reload on GitHub Pages).

## Audience & purpose
Developers and others who spend long hours at a computer.
Goal: explain how to improve their workday for health — eyes, back,
spine, neck.

## Content language
[норвежский / английский — впиши, это важно для всего текста]

## Design rules
- Palette: only three neutral colors — white, black, grey.
- Goal: clean, easy to read, "luxe", no distractions.
- [Tailwind / plain CSS / CSS modules — впиши]
- Mobile-first, responsive.

## Required site structure (from assignment — must match exactly)
1. Risk of health problems
   - Muscle problems
   - Photos/videos of exercises (min. 3, I am in frame)
   - Vision / eyes
   - Psychosocial health problems
2. Requirements for physical organization of the workplace (ergonomics)
   - with specific equipment suggestions
   - Desk / Monitor / Keyboard / Mouse & trackpads (image + text)
3. Links for further exploration  (lives on Home)
[РЕШИ: подразделы = отдельные страницы / якоря / вложенное меню]

## Folder structure
src/
  assets/      # images, svg, content
  components/
  pages/


