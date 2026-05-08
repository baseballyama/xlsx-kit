<script lang="ts">
  import { base } from '$app/paths';
  import { page } from '$app/state';

  const links: Array<{ path: string; label: string; external?: boolean }> = [
    { path: '/docs/getting-started', label: 'Docs' },
    { path: '/api', label: 'API' },
    { path: '/llms.txt', label: 'llms.txt' },
    { path: 'https://github.com/baseballyama/xlsxify', label: 'GitHub', external: true },
  ];

  function resolve(link: (typeof links)[number]): string {
    return link.external ? link.path : `${base}${link.path}`;
  }

  function isActive(link: (typeof links)[number]): boolean {
    if (link.external) return false;
    return page.url.pathname.startsWith(`${base}${link.path}`);
  }
</script>

<header class="site-header">
  <div class="inner">
    <a href="{base}/" class="brand">
      <span class="brand-mark">×</span>
      <span class="brand-name">xlsxify</span>
    </a>
    <nav>
      {#each links as link (link.path)}
        <a href={resolve(link)} class="nav-link" class:active={isActive(link)}>{link.label}</a>
      {/each}
    </nav>
  </div>
</header>

<style>
  .site-header {
    position: sticky;
    top: 0;
    z-index: 10;
    background: color-mix(in oklab, var(--bg) 90%, transparent);
    backdrop-filter: blur(8px);
    border-bottom: 1px solid var(--border);
    height: var(--header-h);
    display: flex;
    align-items: center;
  }

  .inner {
    width: 100%;
    max-width: 1200px;
    margin: 0 auto;
    padding: 0 1.25rem;
    display: flex;
    align-items: center;
    gap: 2rem;
  }

  .brand {
    display: inline-flex;
    align-items: center;
    gap: 0.5rem;
    color: var(--fg);
    font-weight: 700;
    font-size: 1.05rem;
    letter-spacing: -0.01em;
  }

  .brand:hover {
    text-decoration: none;
  }

  .brand-mark {
    display: inline-grid;
    place-items: center;
    width: 28px;
    height: 28px;
    background: var(--accent);
    color: white;
    border-radius: 6px;
    font-size: 1.1rem;
    font-weight: 800;
    line-height: 1;
  }

  nav {
    display: flex;
    gap: 1.25rem;
    align-items: center;
    margin-left: auto;
  }

  .nav-link {
    color: var(--fg-soft);
    font-size: 0.93rem;
    font-weight: 500;
  }

  .nav-link:hover {
    color: var(--fg);
    text-decoration: none;
  }

  .nav-link.active {
    color: var(--fg);
  }
</style>
