export type SupportedAgent = 'cursor' | 'claude' | 'chatgpt' | 'doubao' | 'qwen';

export interface AgentInfo {
  id: SupportedAgent;
  name: string;
  badge: string;
  iconBg: string;
  themeColor: string;
  avatarText: string;
}

export type DemoSceneId = 'scene-august' | 'scene-point-edit' | 'scene-ppt' | 'scene-rollback';

export interface DemoScene {
  id: DemoSceneId;
  title: string;
  subtitle: string;
  description: string;
  userPrompt: string;
  agentActionName: string;
  agentResponse: string;
  toolCallText: string;
  docType: 'excel' | 'ppt';
}
