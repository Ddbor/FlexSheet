/**
 * 内置插件在 Workspace 服务注册表中的键，便于插件间解耦协作。
 */

export const PLUGIN_SERVICE_KEYS = {
  canvas: "flexsheet.service.canvas",
  renderer: "flexsheet.service.renderer",
  theme: "flexsheet.service.theme",
  selection: "flexsheet.service.selection",
  cellEditor: "flexsheet.service.cellEditor",
  flexSheet: "flexsheet.service.flexSheet",
} as const;

export type PluginServiceKey = (typeof PLUGIN_SERVICE_KEYS)[keyof typeof PLUGIN_SERVICE_KEYS];
