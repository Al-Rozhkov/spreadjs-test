<template>
  <div class="app">
    <div class="main">

      <ClientOnly>
        <SpreadSheet />
      </ClientOnly>
    </div>
    <div class="sidebar">...</div>
  </div>
</template>

<style>
.app {
  display: flex;
}

.app .main {
  flex: 1 0 75%;
}
.app .sidebar {
  flex: 1 0 25%;
}
</style>