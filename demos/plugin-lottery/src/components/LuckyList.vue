<script setup lang="ts">
import { state, writeResult } from '../lottery';

const onWriteResult = () => {
  writeResult()
}
</script>
<template>
  <div class="list">
    <table class="list-table">
      <thead>
        <tr>
          <th>奖品类型</th>
          <th>奖品名</th>
          <th>姓名</th>
          <th>工号</th>
        </tr>
      </thead>
      <tbody>
        <tr v-for="item in state.luckyList" :key="item.code">
          <td>{{item.prizeType}}</td>
          <td>{{item.prizeName}}</td>
          <td>{{item.name}}</td>
          <td>{{item.code}}</td>
        </tr>
      </tbody>
    </table>
    <div class="result">
      <button class="result-btn" @click="onWriteResult">写入表格</button>
    </div>
  </div>
  
</template>
<style scoped lang="less">
.list {
  margin-top: 20px; 
  margin-right: 5px;
}
.list-table {
  text-align: center;
  border-collapse: collapse;
  border: 1px solid #ebeef5;
  width: 100%;
  td, th {
    border: 1px solid #ebeef5;
    line-height: 2;
  }

  thead > tr {
    font-size: 14px;
    background-color: #f5f7fa;
  }

  tbody > tr {
    font-size: 14px;
    &:hover {
      background-color: #f5f7fa;
    }
  }
  
}
.result {
  margin-top: 20px;
  text-align: center;
}
.result-btn {
  line-height: 1;
  padding: 0 15px;
  height: 32px;
  outline: 0;
  user-select: none;
  background-color: #409eff;
  border: none;
  border-radius: 4px;
  color: #fff;
  appearance: none;
  cursor: pointer;

  &:hover {
    background-color: #79bbff;
  }
}
</style>