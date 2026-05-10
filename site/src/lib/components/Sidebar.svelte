<script lang="ts">
  import { base } from '$app/paths';
  import { page } from '$app/state';
  import { docSections } from '$lib/docs-nav';
</script>

<aside class="sidebar" data-pagefind-ignore>
  <nav>
    {#each docSections as section, sIdx (section.title)}
      <section>
        <h4>
          <span class="hash">§{sIdx + 1}</span>
          {section.title}
        </h4>
        <ul>
          {#each section.links as link, lIdx (link.href)}
            {@const active = page.url.pathname === `${base}${link.href}`}
            <li>
              <a href="{base}{link.href}" class:active>
                <span class="row-num">{String(lIdx + 1).padStart(2, '0')}</span>
                <span class="row-title">{link.title}</span>
              </a>
            </li>
          {/each}
        </ul>
      </section>
    {/each}
  </nav>
</aside>

<style>
  .sidebar {
    width: var(--sidebar-w);
    flex: 0 0 var(--sidebar-w);
    padding: 2rem 0.75rem 4rem 1.5rem;
    border-right: 1px solid var(--border);
    height: calc(100vh - var(--header-h));
    position: sticky;
    top: var(--header-h);
    overflow-y: auto;
  }

  section {
    margin-bottom: 1.75rem;
  }

  h4 {
    display: flex;
    align-items: baseline;
    gap: 0.55rem;
    font-family: var(--mono);
    font-weight: 600;
    font-size: 11px;
    text-transform: uppercase;
    letter-spacing: 0.14em;
    color: var(--fg-soft);
    margin: 0 0 0.6rem 0.5rem;
    border: none;
    padding: 0;
    font-variation-settings: normal;
  }

  .hash {
    color: var(--accent);
    font-weight: 600;
    font-size: 10px;
    letter-spacing: 0.04em;
  }

  ul {
    list-style: none;
    padding: 0;
    margin: 0;
  }

  li a {
    display: flex;
    align-items: center;
    gap: 0.7rem;
    padding: 0.38rem 0.55rem 0.38rem 0.5rem;
    color: var(--fg-soft);
    border-radius: var(--radius-sm);
    font-size: 0.92rem;
    line-height: 1.35;
    border-left: 2px solid transparent;
  }

  li a:hover {
    color: var(--fg);
    text-decoration: none;
    background: var(--bg-soft);
  }

  li a.active {
    background: var(--accent-soft);
    color: var(--fg);
    border-left-color: var(--accent);
  }

  li a.active .row-num {
    color: var(--accent);
  }

  .row-num {
    flex: none;
    width: 2.2ch;
    text-align: right;
    font-family: var(--mono);
    font-size: 10.5px;
    color: var(--fg-muted);
    font-weight: 500;
    letter-spacing: 0.04em;
  }

  .row-title {
    flex: 1;
  }

  @media (max-width: 800px) {
    .sidebar {
      width: 100%;
      flex: none;
      position: static;
      height: auto;
      border-right: none;
      border-bottom: 1px solid var(--border);
      padding: 1rem;
    }
  }
</style>
