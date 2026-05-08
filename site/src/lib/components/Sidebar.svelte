<script lang="ts">
  import { base } from '$app/paths';
  import { page } from '$app/state';
  import { docSections } from '$lib/docs-nav';
</script>

<aside class="sidebar">
  <nav>
    {#each docSections as section (section.title)}
      <section>
        <h4>{section.title}</h4>
        <ul>
          {#each section.links as link (link.href)}
            <li>
              <a
                href="{base}{link.href}"
                class:active={page.url.pathname === `${base}${link.href}`}
              >
                {link.title}
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
    padding: 1.5rem 0.5rem 4rem 1.25rem;
    border-right: 1px solid var(--border);
    height: calc(100vh - var(--header-h));
    position: sticky;
    top: var(--header-h);
    overflow-y: auto;
  }

  section {
    margin-bottom: 1.5rem;
  }

  h4 {
    font-size: 13px;
    text-transform: uppercase;
    letter-spacing: 0.06em;
    color: var(--fg-muted);
    margin: 0 0 0.5rem 0.5rem;
  }

  ul {
    list-style: none;
    padding: 0;
    margin: 0;
  }

  li a {
    display: block;
    padding: 0.35rem 0.75rem;
    color: var(--fg-soft);
    border-radius: 6px;
    font-size: 0.93rem;
  }

  li a:hover {
    background: var(--bg-soft);
    color: var(--fg);
    text-decoration: none;
  }

  li a.active {
    background: var(--accent-soft);
    color: var(--fg);
    font-weight: 600;
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
