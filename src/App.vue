<template>
  <div id="nav">
    <birthday></birthday>
  </div>
</template>

<script>
import birthday from './components/birthday'
export default {
  name: 'APP',
  components: {
    birthday,
  },
}
</script>

<style lang="less">
html, body{
  height: 100%;
  width: 100%;
  margin: 0;
}
#app {
  font-family: Avenir, Helvetica, Arial, sans-serif;
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
  text-align: center;
  color: #2c3e50;
  min-width: 500px;
}
</style>
