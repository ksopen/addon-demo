<script setup lang="ts">
import { onDrawLottery, state } from '../lottery';

const onLottery = () => {
  onDrawLottery()
}
</script>
<template>
  <div class="lottery-button">
    <button @click="onLottery">开始抽奖</button>
  </div>
  <div class="lottery-info">
    <span v-for="item in state.activePersons">{{item}}</span>
  </div>
  <div class="lottery-tips" v-if="state.activeTimes >= 0 && state.status === 'end'">
    恭喜以上{{state.prizeList[state.activeTimes].prizeNum}}位抽中 <span class="lottery-type">{{state.prizeList[state.activeTimes].prizeType}}</span>，奖品为 <span class="lottery-name">{{state.prizeList[state.activeTimes].prizeName}}</span>
  </div>
</template>
<style scoped>
.lottery-button {
  padding-top: 20px;
  text-align: center;
}

.lottery-info {
  margin-top: 20px;
  font-size: 20px;
  color: red;
  padding: 10px;
  text-align: center;
}
.lottery-info > span {
  padding: 10px;
  margin: 5px 3px;
  display: inline-block;
  border-radius: 5px;
  box-shadow: 0px 0px 3px #afa893;
}

.lottery-tips {
  margin-top: 20px;
  text-align: center;
  color: #417ff9;
}
.lottery-type {
  padding: 3px 5px;
  color: red;
  font-size: 22px;
}
.lottery-name {
  padding: 3px 5px;
  color: red;
  font-size: 22px;
}
button {
  min-width: 150px;
  background: #417ff9;
  color: #fff;
  border-radius: 4px;
  border: none;
  padding: 0.6em 1.2em;
  font-size: 18px;
  font-weight: 500;
  font-family: inherit;
  cursor: pointer;
  transition: all 0.25s;
}
button:hover {
  background-color: #3566ca;
}
</style>