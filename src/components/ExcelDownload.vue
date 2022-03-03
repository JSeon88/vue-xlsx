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
    const data = {
      // 컬럼 넓이(단위 : 노출 글자 수)
      columnWidth: [20, 25, 20, 20, 30, 30, 30, 30],
      merges: [
        {
          start: { rowIndex: 0, columnIndex: 0 },
          end: { rowIndex: 0, columnIndex: 2 },
        },
      ],
      category: [
        [
          {
            text: [
              {
                text: "응시자 일괄 등록 엑셀 서식",
                style: { bold: true, fontSize: 24 },
              },
            ],
          },
        ],
        [
          {
            style: {
              border: true,
              verticalAlignment: "top",
            },
            text: [
              {
                text: "설명",
                style: { bold: true },
              },
            ],
          },
          {
            style: {
              border: true,
              verticalAlignment: "top",
            },
            text: [
              {
                text: "(선택)\n",
              },
              {
                text: "응시자가 속할 그룹명을 입력해주세요.",
              },
            ],
          },
          {
            style: {
              border: true,
              verticalAlignment: "top",
            },
            text: [
              {
                text: "(필수)\n",
                style: { fontColor: "FF0000" },
              },
              {
                text: "응시자 이름을 입력해주세요.",
              },
            ],
          },
          {
            style: {
              border: true,
              verticalAlignment: "top",
            },
            text: [
              {
                text: "(필수)\n",
                style: { fontColor: "FF0000" },
              },
              {
                text: "시험 응시 로그인시 사용할 비밀번호를 입력해주세요.",
              },
            ],
          },
          {
            style: {
              border: true,
              verticalAlignment: "top",
            },
            text: [
              {
                text: "(선택)\n",
              },
              {
                text: "응시자 수집정보 설정에 등록할 내용을 입력해주세요.",
              },
            ],
          },
          {
            style: {
              border: true,
              verticalAlignment: "top",
            },
            text: [
              {
                text: "(선택)\n",
              },
              {
                text: "응시자 수집정보 설정에 등록할 내용을 입력해주세요.",
              },
            ],
          },
          {
            style: {
              border: true,
              verticalAlignment: "top",
            },
            text: [
              {
                text: "(선택)\n",
              },
              {
                text: "응시자 수집정보 설정에 등록할 내용을 입력해주세요.",
              },
            ],
          },
          {
            style: {
              border: true,
              verticalAlignment: "top",
            },
            text: [
              {
                text: "(선택)\n",
              },
              {
                text: "응시자 수집정보 설정에 등록할 내용을 입력해주세요.",
              },
            ],
          },
        ],
        [
          {
            style: {
              border: true,
              verticalAlignment: "top",
            },
            text: [
              {
                text: "예시",
                style: { bold: true },
              },
            ],
          },
          {
            style: {
              border: true,
              verticalAlignment: "top",
            },
            text: [
              {
                text: "아이엠수학학원 3학년 심화반",
                style: { fontColor: "666666" },
              },
            ],
          },
          {
            style: {
              border: true,
              verticalAlignment: "top",
            },
            text: [
              {
                text: "홍길동",
                style: { fontColor: "666666" },
              },
            ],
          },
          {
            style: {
              border: true,
              verticalAlignment: "top",
            },
            text: [
              {
                text: "math03010",
                style: { fontColor: "666666" },
              },
            ],
          },
          {
            style: {
              border: true,
              verticalAlignment: "top",
            },
            text: [
              {
                text: "예)추가정보1 : 휴대전화번호\n010-1234-5678",
                style: { fontColor: "666666" },
              },
            ],
          },
          {
            style: {
              border: true,
              verticalAlignment: "top",
            },
            text: [
              {
                text: "예) 추가정보2 : 학교\n아이엠초등학교",
                style: { fontColor: "666666" },
              },
            ],
          },
          {
            style: {
              border: true,
              verticalAlignment: "top",
            },
            text: [
              {
                text: "예) 추가정보3 : 학년\n3학년",
                style: { fontColor: "666666" },
              },
            ],
          },
          {
            style: {
              border: true,
              verticalAlignment: "top",
            },
            text: [
              {
                text: "예) 추가정보4 : 레벨테스트\nlv.4",
                style: { fontColor: "666666" },
              },
            ],
          },
        ],
        [
          {
            style: {
              border: true,
              fill: "C6C6C6",
            },
            text: " ",
          },
          {
            style: {
              border: true,
              fill: "C6C6C6",
            },
            text: [
              {
                text: "그룹",
                style: { bold: true },
              },
            ],
          },
          {
            style: {
              border: true,
              fill: "C6C6C6",
            },
            text: [
              {
                text: "이름 ",
                style: { bold: true },
              },
              {
                text: "*",
                style: { fontColor: "FF0000" },
              },
            ],
          },
          {
            style: {
              border: true,
              fill: "C6C6C6",
            },
            text: [
              {
                text: "비밀번호 ",
                style: { bold: true },
              },
              {
                text: "*",
                style: { fontColor: "FF0000" },
              },
            ],
          },
          {
            style: {
              border: true,
              fill: "C6C6C6",
            },
            text: [
              {
                text: "추가정보1",
                style: { bold: true },
              },
            ],
          },
          {
            style: {
              border: true,
              fill: "C6C6C6",
            },
            text: [
              {
                text: "추가정보2",
                style: { bold: true },
              },
            ],
          },
          {
            style: {
              border: true,
              fill: "C6C6C6",
            },
            text: [
              {
                text: "추가정보3",
                style: { bold: true },
              },
            ],
          },
          {
            style: {
              border: true,
              fill: "C6C6C6",
            },
            text: [
              {
                text: "추가정보4",
                style: { bold: true },
              },
            ],
          },
        ],
      ],
      list: [
        [
          {
            text: " ",
          },
          {
            text: "EDu",
          },
          {
            text: "Sunny",
          },
          {
            text: "1234",
          },
          {
            text: "추가정보 111",
          },
          {
            text: "추가정보 222",
          },
          {
            text: "추가정보 333",
          },
          {
            text: "추가정보 444",
          },
        ],
      ],
    };

    function excelDownFile() {
      console.log("down...");
      createExcelFile(data, "테스트-다운로드.xlsx");
    }

    return {
      excelDownFile,
    };
  },
};
</script>

<style></style>
