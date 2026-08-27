# SimPy KPI Lab

一个可复现、可扩展的离散事件仿真与安全决策项目：用 SimPy 描述排队系统，用多次 replication 和参数网格做实验，用 KPI 与置信区间比较场景，并可选择调用 OpenAI Responses API 生成结构化分析或白名单动作提案。提案必须经过人工批准，模型没有直接执行 Python 或修改运行中仿真的能力。

## 能力

- 串行多工位 SimPy 模型；每个工位可配置容量和服务时间分布。
- 指数、定值、均匀、三角分布；固定 seed，可复现实验。
- 到达流与各工位使用独立随机流；默认用共同随机数降低场景差值的方差。
- 多次 replication、参数网格、可选多进程并行。
- KPI：到达量、完成量、吞吐率、期末 WIP、平均/P50/P95 周期时间、平均等待、服务水平、工位利用率、工位平均队长与 P95 等待。
- 每项 KPI 输出均值、样本标准差、正态近似置信区间、最小值和最大值。
- JSON 与 CSV 输出；ChatGPT 分析层输出经过 Pydantic 校验的 JSON 和 Markdown。
- OpenAI 安全提案层：只允许 `set_parameter`、`change_policy`、`enable_resource` 三类结构化动作，并再次执行本地 allowlist 校验。
- 人工审批工作流：pending / approved / rejected / applied / failed 状态、乐观版本、幂等 operation ID、配置哈希和有序审计日志。
- FastAPI REST、WebSocket 事件流与可选管理员 Bearer Token；同步仿真和 OpenAI 请求不会阻塞 API 事件循环。
- MCP v2 受限客户端：可校验配置、读取 REST 会话、提交待审批提案和查看事件；刻意不暴露建会话、运行、批准或拒绝工具。
- YAML 严格校验、CLI、pytest、ruff、Dockerfile。

## 快速开始

要求 Python 3.11+。

Windows PowerShell：

```powershell
py -3 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -e ".[dev]"

.\.venv\Scripts\simlab.exe validate examples\service_center.yaml
.\.venv\Scripts\simlab.exe run examples\service_center.yaml --workers 4
.\.venv\Scripts\python.exe -m pytest
```

macOS / Linux：

```bash
python -m venv .venv
source .venv/bin/activate
python -m pip install -e ".[dev]"

simlab validate examples/service_center.yaml
simlab run examples/service_center.yaml --workers 4
pytest
```

若只需要控制 API 或 MCP，也可分别安装：

```powershell
.\.venv\Scripts\python.exe -m pip install -e ".[api]"
.\.venv\Scripts\python.exe -m pip install -e ".[mcp]"
```

运行后会生成：

```text
outputs/service_center/
├── results.json       # 配置、每次 replication、汇总结果
├── replications.csv   # 每次实验的宽表 KPI
└── summary.csv        # 每个场景/指标的统计汇总
```

## 接入 ChatGPT / OpenAI API

API key 只通过环境变量读取，不写入 YAML 或结果文件。

Windows PowerShell：

```powershell
$env:OPENAI_API_KEY = "..."
.\.venv\Scripts\simlab.exe run examples\service_center.yaml --analyze
```

macOS / Linux：

```bash
export OPENAI_API_KEY="..."

# 仿真结束后直接分析
simlab run examples/service_center.yaml --analyze

# 或分析已有结果，并提出业务问题
simlab analyze outputs/service_center/results.json \
  --question "比较各容量方案的服务水平、P95 周期时间和利用率，并指出权衡。"
```

输出 `ai_analysis.json` 与 `ai_analysis.md`。项目使用 Responses API 的 `responses.parse(..., text_format=...)`，让模型输出遵循 Pydantic 数据契约。实现依据：[OpenAI 文本生成指南](https://developers.openai.com/api/docs/guides/text) 与 [Structured Outputs 指南](https://developers.openai.com/api/docs/guides/structured-outputs)。

> AI 分析是决策辅助。KPI 数值与置信区间由本地代码计算；模型只能解释提供的数据，不会改写实验结果。

默认 API 配置显式设置 60 秒超时、2 次重试，并使用 `store: false`。这些值可在 YAML 中覆盖；不带 `--analyze` 的仿真完全不需要 API key 或网络。

## 安全闭环控制 API

这一层实现的是安全的“下一轮实验控制”，不是热修改正在运行的 SimPy generator：

```mermaid
flowchart LR
    A[不可变 ProjectConfig 快照] --> B[运行 SimPy 多次实验]
    B --> C[KPI 汇总]
    C --> D[OpenAI 结构化提案]
    D --> E{人工审批}
    E -->|拒绝| F[审计并结束]
    E -->|批准| G[本地 allowlist + 完整配置校验]
    G --> H[新配置版本]
    H --> A
```

默认仅监听本机。建议先设置一个管理 Token：

```powershell
$env:SIMLAB_API_TOKEN = "请替换为随机长字符串"
$env:SIMLAB_API_OUTPUT_ROOT = "outputs/api_sessions"
.\.venv\Scripts\simlab.exe serve --host 127.0.0.1 --port 8000
```

打开 `http://127.0.0.1:8000/docs` 可直接使用 OpenAPI 页面。PowerShell 完整演示见 `examples/human_in_the_loop.ps1`：

```powershell
.\examples\human_in_the_loop.ps1
```

核心 REST 路由：

- `POST /v1/sessions`：创建配置会话。
- `GET /v1/sessions/{id}/allowed-actions`：读取当前精确动作白名单。
- `POST /v1/sessions/{id}/proposals`：人工或其他可信系统提交结构化提案。
- `POST /v1/sessions/{id}/proposals:generate`：让 OpenAI 根据最近一次 KPI 汇总生成待审批提案。
- `POST /v1/sessions/{id}/proposals/{proposal_id}:approve`：人工批准并应用到下一版配置。
- `POST /v1/sessions/{id}/proposals/{proposal_id}:reject`：人工拒绝。
- `POST /v1/sessions/{id}/runs`：以当前不可变配置快照运行实验。
- `GET /v1/sessions/{id}/audit`：读取审批审计与服务事件。
- `WS /v1/sessions/{id}/events?after_sequence=0`：用 `Authorization: Bearer ...` 请求头按序订阅或重放事件，Token 不放进 URL。

每个写操作都要求调用方提供唯一的 `operation_id` 和当前 `expected_version`。重复的同一操作会返回原结果；复用 ID 发送不同内容或使用旧版本会得到冲突，而不是覆盖较新的状态。

服务端不会使用请求中的 `experiment.output_dir` 写文件。每次运行都写到服务端分配的 `输出根目录/会话 ID/run ID/`，避免并发运行相互覆盖或路径穿越。

### OpenAI 动作安全边界

OpenAI 使用 Responses API 的 Pydantic Structured Outputs 生成提案。发送给模型的内容仅包括当前配置、KPI 汇总、业务目标和允许动作目录。返回后仍会执行以下检查：

1. 动作类型与 target 必须逐字命中本地 allowlist。
2. `requires_approval` 必须为 `true`。
3. 动作应用到配置副本后，必须重新通过完整 `ProjectConfig` 校验；每个参数网格场景还要分别检查时长、工位数、单次到达量和总事件预算，不能用 grid 绕过基础上限。
4. 提案的配置哈希必须仍与当前版本一致，旧提案不能套用到新配置。
5. 模型不能设置 API key、模型名、输出路径、随机 seed 或任意 dotted path，也没有 shell、文件、Python、`eval` 或 `exec` 工具。

一个 AI plan 可以包含多个有顺序的动作。每一步分别人工审批，后一步绑定前一步应用后的配置哈希；跳过或拒绝前置动作时，后续动作会按过期提案拒绝，需要基于当前配置重新生成。它不是“一次批准全部”的原子事务。为避免停用工位后列表下标改变，`enable_resource` 必须排在所有 `simulation.stations.<index>` 参数动作之后；服务端也会强制检查这个顺序。

Responses API 支持自定义函数调用和结构化 JSON 输出；本项目选择更严格的“只生成数据提案、由本地代码执行”边界。实现依据：[OpenAI Responses API](https://developers.openai.com/api/reference/python/resources/responses/methods/create) 与 [Structured Outputs](https://developers.openai.com/api/docs/guides/structured-outputs)。

## MCP 接口

安装 `[mcp]` 后启动本地 stdio 服务：

```powershell
$env:SIMLAB_API_URL = "http://127.0.0.1:8000"
$env:SIMLAB_API_TOKEN = "与 REST API 相同的管理 Token"
.\.venv\Scripts\simlab-mcp.exe
```

独立 MCP 进程通过 `SIMLAB_API_URL` 连接正在运行的 REST 服务，因此两端读取的是同一会话、版本和审计记录。可用工具仅包括 `validate_project`、`get_session`、`get_allowed_actions`、`submit_proposal` 和 `list_events`。MCP 不暴露建会话、运行、批准或拒绝；会话创建、仿真启动与人工决策必须在 REST/OpenAPI 界面完成，模型不能用“另建会话再运行”绕过审批边界。实现使用 MCP Python SDK v2，本地宿主默认采用 stdio。

## 当前控制层限制

- 会话、幂等记录和事件日志目前保存在单进程内存中；服务重启后会丢失。生产部署应换成 SQLite/PostgreSQL，并在持久化后再广播事件。
- WebSocket broker 面向单个 Uvicorn worker；多实例部署应使用 Redis 或其他消息总线。
- 已运行的 SimPy replication 不会暂停、序列化或热改容量。批准动作只影响下一次运行，这保证 seed、配置和结果能够审计复现。
- API 暂不提供角色型 JWT；设置 `SIMLAB_API_TOKEN` 时使用单一管理员 Bearer Token。公网部署前应接入正式身份和 approver 角色。

## 配置说明

最小配置：

```yaml
simulation:
  until: 480
  warmup: 60
  arrival_interarrival: {kind: exponential, mean: 5}
  stations:
    - name: service
      capacity: 2
      service_time: {kind: triangular, low: 4, mode: 8, high: 15}

experiment:
  replications: 20
  base_seed: 20260825
  common_random_numbers: true
  parameter_grid:
    stations.0.capacity: [1, 2, 3]

openai:
  model: gpt-5.6
  max_output_tokens: 2500
  timeout_seconds: 60
  max_retries: 2
  store: false
```

`parameter_grid` 使用相对 `simulation` 的点路径；列表下标也写在路径中。例如：

```yaml
parameter_grid:
  stations.0.capacity: [1, 2]
  arrival_interarrival.mean: [4.0, 5.0, 6.0]
```

会生成 2 × 3 = 6 个场景，每个场景再执行 `replications` 次。也支持以 `simulation.` 开头的完整路径。

`common_random_numbers: true` 会让不同场景的同编号 replication 使用相同 seed，并由基于 BLAKE2b 稳定派生的命名随机流分别驱动到达和各工位服务时间，适合做容量方案的配对比较。工位随机流按名称派生，因此新增或调整其他工位不会意外改变已有工位的随机序列。若场景结构差异很大、不希望随机输入配对，可设为 `false`。

分布格式：

```yaml
# 指数分布
{kind: exponential, mean: 5}

# 定值
{kind: deterministic, value: 5}

# 均匀分布
{kind: uniform, low: 3, high: 7}

# 三角分布
{kind: triangular, low: 3, mode: 5, high: 9}
```

所有时间量使用同一个、由业务自行定义的时间单位；示例使用“分钟”。

## KPI 口径

- 统计窗口为 `[warmup, until)`。
- `arrivals`：统计窗口内的到达数。
- `completed` / `throughput_per_time_unit`：窗口内完成的实体数及其除以窗口长度。
- 周期时间仅统计在 warmup 后到达、且在仿真结束前完成的实体，避免 warmup 前队列污染 cohort 指标。
- 工位等待时间仅统计 warmup 后进入该工位队列、并在结束前开始服务的实体。
- 利用率按统计窗口内实际服务区间 /（capacity × 窗口长度）计算；跨边界的服务区间自动裁剪。
- 平均队长由每个实体的排队时间面积计算，仿真结束时仍在排队的实体也计入。
- `service_level` 为周期时间不超过 `cycle_time_target` 的 cohort 比例。
- `avg_wait_time` 是所有有效工位访问的平均单次等待；`avg_total_wait_time` 是已完成 cohort 每个实体跨工位的平均总等待。
- `censored_cycle_count` 与 `cycle_completion_fraction` 显示仿真结束时的右删失程度，避免只看已完成实体而忽略尾部积压。
- 汇总表按“主要 KPI / 驱动指标 / 护栏 / 数据质量”标注角色、方向和单位；容量、样本数等元数据不会混入 KPI 汇总。
- 置信区间使用 replication 均值的正态近似；`n_total`、`n_missing` 会显式显示缺失值，只有一次有效 replication 时标准差和置信区间为 `null`，不会伪装成零不确定性。replication 较少或重尾数据明显时，建议增加次数或改为 bootstrap / t 区间。

## 如何扩展模型

当前示例是“所有实体依次经过所有工位”的串行网络。常见扩展点：

1. 在 `simulation.py` 的 `customer()` 中加入条件路由、返工、优先级或批处理。
2. 在 `config.py` 中新增业务参数，保持 `extra="forbid"` 以尽早发现拼写错误。
3. 在 `KPICollector` 中通过事件记录新增 KPI；不要让 AI 参与 KPI 数值计算。
4. 为新增的随机过程延续“独立随机流”规则，避免改变某个工位后意外改变到达序列。
5. 为定制模型增加一个新的 runner，而不是把业务规则硬编码进 CLI。

## 设计结构

```mermaid
flowchart TD
    A[YAML / JSON 配置] --> B[场景参数网格]
    B --> C[SimPy replications]
    C --> D[KPI 原始记录]
    D --> E[统计汇总与置信区间]
    E --> F[JSON / CSV]
    E --> G[OpenAI 结构化分析]
    E --> H[OpenAI 白名单动作提案]
    H --> I[版本化人工审批]
    I --> J[下一轮不可变配置]
    J --> B
```

OpenAI 层是可选适配器；没有 API key 时，仿真、KPI 与导出功能仍可完整运行。

## Docker

```bash
docker build -t simpy-kpi-lab .
docker run --rm -v "$PWD/outputs:/app/outputs" simpy-kpi-lab
```

如需 API 分析：

```bash
docker run --rm \
  -e OPENAI_API_KEY \
  -v "$PWD/outputs:/app/outputs" \
  simpy-kpi-lab run examples/service_center.yaml --analyze
```
