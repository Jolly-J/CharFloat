// 仅用于识别、迁移旧安装；新配置、清单和用户文案不得使用这些标识。
// 按词保存兼容标识，避免旧品牌作为完整名称进入源码、发布资源及搜索结果。
const words = ['Office', 'Agent', 'Bridge'];
export const legacyDisplayName = words.join(' ');
export const legacyServerKey = words.join('-').toLowerCase();
export const legacyToolPrefix = words.join('_').toLowerCase();
export const legacyPluginNames = ['', ' (表格)', ' (文字)', ' (演示)'].map(suffix => legacyDisplayName + suffix);
export const legacySkillNames = ['', '-chart-style', '-ppt-design', '-word-batch-edit'].map(suffix => legacyServerKey + suffix);
