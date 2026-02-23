# JustDiff / ExcelMerge 优化与改进建议

## 一、必须修复的 Bug

### 1. xlsmerge 双地址对比时第一个 URI 传错（Entrance.cs）

**位置**：`ProcessInput` 中解析 `&cmp=` 分支。

**问题**：`DiffUri(rev, new Uri(url), cmprev, new Uri(cmpfileurl))` 里第一个参数用了 `url`（仍包含 `&cmp=...` 的完整字符串），应使用解析后的 `fileurl`。

**修改**：
```csharp
// 原：DiffUri(rev, new Uri(url), cmprev, new Uri(cmpfileurl));
DiffUri(rev, new Uri(fileurl), cmprev, new Uri(cmpfileurl));
```

### 2. Config.EmptyLine 被就地修改（Util.cs WorkBookWrap.CalValideRow）

**位置**：`CalValideRow` 内 `if (cfg.EmptyLine-- > 0)`。

**问题**：直接对 `config` 做 `EmptyLine--`，加载第一个工作簿会改掉全局配置，加载第二个工作簿时用的已是递减后的值，两个工作簿行为不一致。

**建议**：在方法内用局部变量，例如：
```csharp
int emptyLine = cfg.EmptyLine;
// ...
if (emptyLine-- > 0) { continue; }
else { ... break; }
```

---

## 二、异常与错误处理

### 1. 吞掉异常的 catch（MainWindow 构造函数）

**位置**：注册 `xlsmerge` 协议时的 `catch { }`。

**问题**：无日志、无提示，权限不足或注册失败时难以排查。

**建议**：至少记录日志或仅在 Debug 下抛出，例如：
```csharp
catch (Exception ex) {
    System.Diagnostics.Debug.WriteLine($"xlsmerge 协议注册失败: {ex.Message}");
}
```

### 2. SVN / 文件操作未 try-catch

**位置**：
- `Entrance.GetVersionFile`：`client.Write`、`File.Create` 可能抛异常。
- `Entrance.DiffUri` / `Diff`：网络或本地路径错误会直接崩溃。
- `MainWindow.Refresh`：`new WorkBookWrap(file, config)` 里若文件损坏或格式不对会抛。

**建议**：在入口处（如 `ProcessInput`、`OnDragFile`、`Refresh`）对 SVN 与文件 IO 做 try-catch，向用户提示“无法打开/对比该文件”并记录异常信息，避免整个进程退出。

### 3. getUrl 中 long.Parse 可能抛

**位置**：`getUrl` 内 `rev = long.Parse(srev)`。

**问题**：URL 中版本号非数字时会 FormatException。

**建议**：使用 `long.TryParse(srev, out rev)`，解析失败时用 0 或返回 false 由调用方处理。

---

## 三、资源与生命周期

### 1. IWorkbook 未释放

**位置**：`WorkBookWrap` 持有一个 `IWorkbook book`（NPOI），当前未实现 `IDisposable`，也未在窗口关闭时关闭工作簿。

**问题**：大文件或多次切换文件可能导致句柄/内存占用偏高。

**建议**：  
- 为 `WorkBookWrap` 实现 `IDisposable`，在 `Dispose` 中关闭 `book`（如 NPOI 的 Close/Dispose）。  
- 在 `MainWindow` 关闭或下次 `Refresh` 替换 books 前，对旧的 `WorkBookWrap` 调用 `Dispose`。

### 2. 临时 SVN 文件仅在 Window_Closing 清理

**位置**：`Entrance._tempFiles` 在 `Window_Closing` 里统一删。

**问题**：若用户从未关闭主窗口就退出进程，或通过任务管理器结束，临时文件可能残留。

**建议**：  
- 使用 `Path.GetTempFileName()` 或带前缀的临时路径，便于识别与批量清理。  
- 进程退出时再扫一次临时目录清理本进程创建的文件（例如在 App.Exit 或析构中），或记录到配置文件在下次启动时清理。

---

## 四、可读性与可维护性

### 1. 魔法字符串与硬编码

**位置**：  
- `Entrance.DiffList`：`"http://m1.svn.ejoy.com/m1/" + file` 写死。  
- 多处 `"src"` / `"dst"` 字符串。

**建议**：  
- SVN 根 URL 放到 `config.json` 或配置类（如 `Config.SvnBaseUrl`），`DiffList` 从配置读取。  
- 对 `"src"`/`"dst"` 可考虑常量或枚举，减少拼写错误。

### 2. getUrl 命名与返回值

**位置**：`getUrl(path, out string url, out long rev)`。

**问题**：`url` 实际是“整段 path”（含查询串），易误解；且解析失败时无明确信号。

**建议**：  
- 改为 `ParseRevisionFromUrl(string path, out string baseUrl, out long revision)`，或返回 `(string baseUrl, long revision)?`，解析失败返回 null。  
- 若保留 out，则 `baseUrl` 建议去掉 `?r=...` 等查询部分，供 `new Uri(baseUrl)` 使用，避免歧义。

### 3. 未使用的 using

**位置**：`Entrance.cs` 首行 `using NPOI.Util;`。

**建议**：删除未使用的 using，保持文件简洁。

---

## 五、线程与 UI

### 1. 大文件/大目录在 UI 线程做 Diff

**位置**：`MainWindow.Refresh`、`YAMLDifferWindow.Refresh`、目录对比等都在主线程执行。

**问题**：大 Excel 或大 YAML 会卡住界面。

**建议**：  
- 将“加载 + 计算 diff”放到 `Task.Run`，完成后再 `Dispatcher.Invoke` 更新 UI。  
- 在计算期间显示“加载中/对比中”并禁用相关按钮，避免重复点击。

### 2. YAMLDifferWindow 中 Dispatcher.BeginInvoke

**位置**：`Refresh()` 末尾 `Dispatcher.BeginInvoke(()=>{ SrcDataGrid.ExpandAll(); ... });`

**说明**：若 `RefreshData()` 已保证在 UI 线程完成，则 ExpandAll 用 `BeginInvoke` 可接受；若 `RefreshData` 在后台线程，需确保 UI 更新都在主线程执行。

---

## 六、配置与兼容性

### 1. config.json 缺失字段

**位置**：反序列化 `Config` 时，若旧版 config 缺少新字段，会使用 C# 默认值。

**建议**：  
- 保持默认值合理（当前类里已对部分字段赋初值）。  
- 若以后增加必填项，可考虑版本号或迁移逻辑，避免旧配置导致异常。

### 2. 注册表协议与权限

**位置**：`Registry.ClassesRoot.CreateSubKey(@"xlsmerge\...")` 需要管理员权限。

**建议**：  
- 安装/首次运行时用 manifest 或安装器请求管理员，或改为 HKCU 下注册（仅当前用户），避免静默失败。  
- 当前 catch 至少不要完全吞掉异常（见上文）。

---

## 七、小结优先级

| 优先级 | 项 |
|--------|----|
| 高 | 修复 xlsmerge 双地址时第一个 URI 使用 `fileurl`；修复 `Config.EmptyLine` 被就地修改 |
| 中 | SVN/文件异常处理与用户提示；GetVersionFile / getUrl 的异常与解析失败处理；IWorkbook 释放与临时文件清理策略 |
| 低 | 去掉未使用 using；魔法字符串进配置；getUrl 重命名与返回值；大文件时异步 diff + 加载中状态 |

如需，我可以按上述顺序给出对应文件的具体补丁（diff）或分步修改说明。
