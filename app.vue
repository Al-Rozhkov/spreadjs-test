<template>
  <div class="app">
    <ul class="menu">
      <li>
        <NuxtLink to="/">Front page</NuxtLink>
      </li>
      <li>
        <NuxtLink to="/outline">Outline</NuxtLink>
      </li>
      <li>
        <NuxtLink to="/charts">Charts</NuxtLink>
      </li>
      <li>
        <NuxtLink to="/pivot-tables">Pivot Tables</NuxtLink>
      </li>
    </ul>

    <NuxtPage />
  </div>
</template>

<style lang="scss">
.menu {
  list-style-type: none;
  margin: 1rem 0 1rem;
  display: flex;
  gap: 2rem;
}
</style>