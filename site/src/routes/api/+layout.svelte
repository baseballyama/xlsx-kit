<script lang="ts">
  import { base } from '$app/paths';
  import { page } from '$app/state';
  import type { LayoutProps } from './$types';

  const { data, children }: LayoutProps = $props();

  const currentSection = $derived(page.params.section ?? null);
</script>

<div class="layout">
  <aside class="sidebar">
    <h4>API reference</h4>
    <ul>
      <li>
        <a href="{base}/api" class:active={currentSection === null}>Overview</a>
      </li>
    </ul>
    <ul>
      {#each data.sections as section (section.id)}
        <li>
          <a
            href="{base}/api/{section.id}"
            class:active={currentSection === section.id}
          >
            <span class="title">{section.title}</span>
            <span class="count">{section.itemCount}</span>
          </a>
        </li>
      {/each}
    </ul>
  </aside>
  <div class="main">
    {@render children?.()}
  </div>
</div>

<style>
  .layout {
    display: flex;
    align-items: stretch;
    max-width: 1500px;
    margin: 0 auto;
  }

  .sidebar {
    width: 280px;
    flex: 0 0 280px;
    padding: 1.5rem 0.5rem 4rem 1.25rem;
    border-right: 1px solid var(--border);
    height: calc(100vh - var(--header-h));
    position: sticky;
    top: var(--header-h);
    overflow-y: auto;
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
    margin: 0 0 1rem 0;
  }

  li a {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 0.5rem;
    padding: 0.4rem 0.7rem;
    color: var(--fg-soft);
    border-radius: 6px;
    font-size: 0.9rem;
    line-height: 1.3;
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

  .title {
    flex: 1;
    min-width: 0;
  }

  .count {
    font-size: 13px;
    color: var(--fg-muted);
    background: var(--bg-soft);
    border: 1px solid var(--border);
    padding: 0.05em 0.45em;
    border-radius: 999px;
    font-variant-numeric: tabular-nums;
  }

  .main {
    flex: 1;
    min-width: 0;
    padding: 0 1rem;
  }

  @media (max-width: 800px) {
    .layout {
      flex-direction: column;
    }
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
