<script setup lang="ts">
import { reactive } from 'vue'
import Lottery from './components/Lottery.vue'
import LuckyList from './components/LuckyList.vue'
const state = reactive({
  activeItem: 'lottery'
})

const onChangeItem = (active: string) => {
  state.activeItem = active
}
</script>

<template>
  <div class="lottery-wrap">
    <div class="lottery-items">
      <div 
        class="lottery-item" 
        :class="{
          'lottery-item-active': state.activeItem === 'lottery'
        }" 
        @click="onChangeItem('lottery')">抽奖</div>
      <div 
        class="lottery-item" 
        :class="{
          'lottery-item-active': state.activeItem === 'list'
        }" @click="onChangeItem('list')">中奖名单</div>
    </div>
    <div class="lottery-content">
      <Lottery v-if="state.activeItem === 'lottery'"/>
      <LuckyList v-if="state.activeItem === 'list'"/>
    </div>
  </div>
</template>

<style scoped>

.lottery-wrap {
  display: flex;
  flex-direction: column;
  padding-top: 40px;
  width: 100%;
  height: 100%;
  box-sizing: border-box;
}
.lottery-items {
  position: fixed;
  width: 100%;
  top: 0;
  left: 0;
  display: flex;
  flex: 0;
  border-bottom: 1px solid #e7e9eb;
  z-index: 10;
  background-color: #fff;
}

.lottery-item {
  position: relative;
  top: 1px;
  padding: 6px 10px;
  margin-right: 10px;
  font-size: 14px;
  border: 1px solid #e7e9eb;
  border-bottom: none;
  border-radius: 3px;
  cursor: pointer;
}
.lottery-item-active {
  color: #417ff9;
  border-bottom: 1px solid #fff;
}

.lottery-tips {
  position: fixed;
  bottom: 0px;
  left: 10px;
  font-size: 12px;
  line-height: 16px;
  color: #999;
}
.lottery-content {
  flex: 1;
  overflow-y: auto;
}
</style>
<style>
::-webkit-scrollbar {
    width: 8px;
    height: 8px;
}
::-webkit-scrollbar-thumb {
  background-color: #e7e9eb;
  border-radius: 4px;
}
::-webkit-scrollbar-track {
  background-color: transparent;
}
</style>