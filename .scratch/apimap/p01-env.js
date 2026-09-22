// 环境基本信息（只读）
const out = { component: String(params && params.component || 'unknown') };
try { out.appName = app.Name; } catch (e) { out.appNameErr = String(e); }
try { out.appVersion = app.Version; } catch (e) { out.appVersionErr = String(e); }
try { out.appBuild = app.Build; } catch (e) { out.appBuildErr = String(e); }
try { out.appPath = app.Path; } catch (e) { out.appPathErr = String(e); }
try { out.operationSystem = app.OperatingSystem; } catch (e) { out.osErr = String(e); }
try { out.wpsGlobal = typeof wps; } catch (e) { out.wpsGlobalErr = String(e); }
const keys = [];
try { for (const k in app) { keys.push(k); } } catch (e) { out.forInErr = String(e); }
out.appForInCount = keys.length;
out.appForInSample = keys.slice(0, 120);
try { out.appOwnNames = Object.getOwnPropertyNames(app).slice(0, 200); } catch (e) { out.appOwnErr = String(e); }
return out;
