import type { DropdownItem } from "../toolbar/toolbar-dropdown.js";

/**
 * 开始选项卡「字体」下拉：顺序与 Excel 常见列表一致，每项以对应字体展示名称。
 */
export const RIBBON_FONT_FAMILY_ITEMS: readonly DropdownItem[] = [
  {
    id: "home.font.family.msYahei",
    label: "微软雅黑",
    previewFontFamily: '"Microsoft YaHei", "微软雅黑", "PingFang SC", sans-serif',
  },
  {
    id: "home.font.family.simHei",
    label: "黑体",
    previewFontFamily: 'SimHei, "Heiti SC", "STHeiti", sans-serif',
  },
  {
    id: "home.font.family.nsimsun",
    label: "新宋体",
    previewFontFamily: 'NSimSun, SimSun, "Songti SC", serif',
  },
  {
    id: "home.font.family.fangSong",
    label: "仿宋",
    previewFontFamily: 'FangSong, "STFangsong", "FangSong_GB2312", serif',
  },
  {
    id: "home.font.family.liSu",
    label: "隶书",
    previewFontFamily: 'LiSu, "Baoli SC", serif',
  },
  {
    id: "home.font.family.kaiTi",
    label: "楷体",
    previewFontFamily: 'KaiTi, "Kaiti SC", "STKaiti", serif',
  },
  {
    id: "home.font.family.arial",
    label: "Arial",
    previewFontFamily: "Arial, sans-serif",
  },
  {
    id: "home.font.family.arialBlack",
    label: "Arial Black",
    previewFontFamily: '"Arial Black", Arial, sans-serif',
  },
  {
    id: "home.font.family.calibri",
    label: "Calibri",
    previewFontFamily: "Calibri, sans-serif",
  },
  {
    id: "home.font.family.cambria",
    label: "Cambria",
    previewFontFamily: "Cambria, serif",
  },
  {
    id: "home.font.family.candara",
    label: "Candara",
    previewFontFamily: "Candara, sans-serif",
  },
  {
    id: "home.font.family.century",
    label: "Century",
    previewFontFamily: "Century, serif",
  },
  {
    id: "home.font.family.courierNew",
    label: "Courier New",
    previewFontFamily: '"Courier New", Courier, monospace',
  },
  {
    id: "home.font.family.comicSansMs",
    label: "Comic Sans MS",
    previewFontFamily: '"Comic Sans MS", cursive, sans-serif',
  },
  {
    id: "home.font.family.garamond",
    label: "Garamond",
    previewFontFamily: "Garamond, serif",
  },
  {
    id: "home.font.family.georgia",
    label: "Georgia",
    previewFontFamily: "Georgia, serif",
  },
  {
    id: "home.font.family.malgunGothic",
    label: "Malgun Gothic",
    previewFontFamily: '"Malgun Gothic", sans-serif',
  },
  {
    id: "home.font.family.mangal",
    label: "Mangal",
    previewFontFamily: "Mangal, sans-serif",
  },
  {
    id: "home.font.family.meiryo",
    label: "Meiryo",
    previewFontFamily: "Meiryo, sans-serif",
  },
  {
    id: "home.font.family.msGothic",
    label: "MS Gothic",
    previewFontFamily: '"MS Gothic", monospace',
  },
  {
    id: "home.font.family.msMincho",
    label: "MS Mincho",
    previewFontFamily: '"MS Mincho", serif',
  },
  {
    id: "home.font.family.msPGothic",
    label: "MS PGothic",
    previewFontFamily: '"MS PGothic", sans-serif',
  },
  {
    id: "home.font.family.msPMincho",
    label: "MS PMincho",
    previewFontFamily: '"MS PMincho", serif',
  },
  {
    id: "home.font.family.tahoma",
    label: "Tahoma",
    previewFontFamily: "Tahoma, sans-serif",
  },
  {
    id: "home.font.family.times",
    label: "Times",
    previewFontFamily: "Times, 'Times New Roman', serif",
  },
  {
    id: "home.font.family.timesNewRoman",
    label: "Times New Roman",
    previewFontFamily: '"Times New Roman", Times, serif',
  },
  {
    id: "home.font.family.trebuchetMs",
    label: "Trebuchet MS",
    previewFontFamily: '"Trebuchet MS", sans-serif',
  },
  {
    id: "home.font.family.verdana",
    label: "Verdana",
    previewFontFamily: "Verdana, sans-serif",
  },
  /* 符号字体：用自身 font-family 渲染拉丁字母会映射成图标，菜单里会像乱码，故不设置 preview */
  {
    id: "home.font.family.wingdings",
    label: "Wingdings",
  },
];

export const RIBBON_FONT_FAMILY_DEFAULT_PREVIEW =
  '"Microsoft YaHei", "微软雅黑", "PingFang SC", sans-serif';
