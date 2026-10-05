# dsh-tool-microsoft-365

[English](README.md) | [中文](README.zh.md)

这是一个面向 [DeepSeek Harness](https://github.com/deepseek-ai/deepseek-harness) 的 Microsoft 365 只读插件。插件通过 Microsoft Graph v1.0 REST API 查询当前登录用户的资料、Outlook 邮件、日历、OneDrive 元数据和 Teams 聊天搜索结果。

插件不会发送邮件、修改日程、上传或下载文件、修改 Teams 数据，也不会调用 Microsoft Graph 写入接口。

## 安装

~~~sh
npm install @libai168/dsh-tool-microsoft-365
~~~

需要 peer dependency：@deepseek-ai/cordis (^4.0.1) 和 @deepseek-ai/dsh-tools (^0.1.0-rc.6)。

## 认证与配置

由宿主配置传入 Microsoft Graph 委托权限 OAuth access token。请将 token 放在密钥管理系统中，不要提交到代码库；任何工具都不会返回 token。

~~~yaml
- name: '@libai168/dsh-tool-microsoft-365'
  config:
    accessToken: 'replace-with-a-short-lived-delegated-graph-token'
    # timeoutMs: 15000
~~~

完整示例见 [examples/cordis.yml](examples/cordis.yml)。客户端只接受生产端点 https://graph.microsoft.com/v1.0，防止 Bearer token 被发送到任意主机。OAuth 获取和刷新由宿主负责。

建议只申请启用工具所需的最小委托权限：

| 能力 | Microsoft Graph 委托权限 |
|---|---|
| ms_auth_test、ms_get_profile | User.Read |
| ms_search_messages | Mail.Read |
| ms_list_calendar_events、ms_list_calendar_view | Calendars.Read |
| ms_search_drive_items、ms_get_drive_item | Files.Read |
| ms_search_chat_messages | Chat.Read |

Microsoft Entra 同意流程和 Conditional Access 策略可能要求管理员批准。权限不足时，工具只返回 Graph 状态和经过清理的错误码。

## 工具

| 工具 | 说明 | 权限类型 |
|---|---|---|
| ms_auth_test | 通过 /me 验证委托 token，并返回安全的资料字段 | read |
| ms_get_profile | 读取当前用户的资料白名单字段 | read |
| ms_search_messages | 通过 POST /search/query 搜索 Outlook 邮件 | read |
| ms_list_calendar_events | 列出默认日历或指定日历的日程 | read |
| ms_list_calendar_view | 在 ISO 8601 时间窗口中列出日程实例 | read |
| ms_search_drive_items | 搜索 OneDrive 文件和文件夹 | read |
| ms_get_drive_item | 读取单个 OneDrive 项目的安全元数据 | read |
| ms_search_chat_messages | 通过 POST /search/query 搜索 Teams 聊天消息 | read |

邮件和 Teams 搜索使用 Microsoft Graph 只读 Search API，实体类型分别为 message 和 chatMessage。日历和 OneDrive 集合分页在插件内部处理 Graph 的 @odata.nextLink。

## 分页

列表工具在 Graph 有下一页时返回短的进程内 nextCursor。将它作为 cursor 传给同一个工具即可继续读取。cursor 使用随机值，不包含 Graph URL 或 skip token；插件进程退出后失效。插件会检查每个 cursor 是否属于当前工具固定的资源路径。

搜索工具使用 offset 分页，返回 nextFrom，下一次请求将它作为 from 传入。offset 和每页数量都有上限，避免无意中请求无限结果。

## 安全与数据边界

- 只接受 https://graph.microsoft.com/v1.0；任意 HTTP 主机和绝对 cursor URL 都会在发送请求前拒绝。
- 每次请求使用 Authorization: Bearer；token 和原始请求头不会出现在工具结果中。
- 不把 Graph 原始错误消息复制给模型，只返回 HTTP 状态、经过校验的短错误码和截断的 request ID。
- 资料姓名、UPN/邮箱、组织字段、发件人字段、主题、地点和 URL 都有限长。Outlook 和 Teams 内容限制为 1000 字符预览，Teams HTML 会先移除标签。
- OneDrive 结果不包含文件内容和预授权下载 URL，只返回必要元数据。
- 请求支持 timeoutMs 和工具执行 AbortSignal。HTTP 429 和服务端错误会抛回宿主，由宿主统一重试。

## 限制

- MVP 只支持委托 token；应用权限凭据和 OAuth 流程由宿主处理。
- 搜索分页使用 Graph offset 语义，邮箱或聊天索引变化可能影响连续请求结果。
- 日历和 OneDrive cursor 只在当前进程有效，重启后需要从第一页开始。
- 没有发送、写入、删除、上传、下载、订阅或管理类工具。

## 开发

~~~sh
npm install
npm run typecheck
npm run build
npm pack --dry-run
~~~

端点映射和发布检查见 [DEVELOPMENT.md](DEVELOPMENT.md)。

## 许可证

[MIT](LICENSE)
