<template>
  <div>
    <button @click.prevent="excelDownFile">다운로드</button>
  </div>
</template>

<script>
import { createExcelFile } from "../common/excel.js";

export default {
  name: "ExcelDownload",
  setup() {
    // 서버에서 추가정보 리스트 가져오기
    const rawData = [
      [
        [
          {
            text: "a1-1\n",
            style: { italic: true, bold: true },
          },
          {
            text: "a1-2",
            style: { fontColor: "FF0000" },
          },
        ],
        null,
        [
          {
            text: "c1-1\n",
            style: { italic: true, bold: true },
          },
          {
            text: "c1-2",
            style: { fontColor: "FF0000" },
          },
        ],
      ],
      [
        [
          {
            text: "a2-1\n",
            style: { italic: true, bold: true },
          },
          {
            text: "a2-2",
            style: { fontColor: "FF0000" },
          },
        ],
        "b2",
        [
          {
            text: "c2-1\n",
            style: { italic: true, bold: true },
          },
          {
            text: "c2-2",
            style: { fontColor: "FF0000" },
          },
        ],
      ],
    ];

    function excelDownFile() {
      console.log("down...");
      createExcelFile(rawData, "테스트-다운로드.xlsx");
    }

    return {
      excelDownFile,
    };
  },
};
</script>

<style></style>
