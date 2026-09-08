# 上游同步信息

本目录（`vendor/core/`）是 [MarsandSea/excel-router](https://github.com/MarsandSea/excel-router)
仓库 `core/` 目录的逐字副本，由 `.github/workflows/sync-upstream.yml` 自动维护。

**不要手动修改 `vendor/core/` 下的任何文件** —— 下次自动同步会直接覆盖，改动会丢失。
如果发现 vendor 代码有 bug，应该去上游仓修，等下一次同步带过来；紧急情况下可以在
`scripts/` 下的 CLI 包装层里用 monkeypatch 打补丁（参考 `er_pdf_dist.py` 覆盖
`_find_cjk_font` 的做法），不要碰 vendor 本身。

- **上游仓库**：https://github.com/MarsandSea/excel-router
- **同步依据**：上游最新的 `v*` 发布 tag（不跟 main 分支的半成品，避免把没测试过的改动带进 skill）
- **当前记录的 tag**：见同目录 `.upstream-tag`（单行纯文本，供同步 Action 做单调递增比较，不要手改）

<!-- SYNC:BEGIN 以下内容由 sync-upstream.yml 每次同步自动重写，手改无效 -->
## 当前同步状态

- 上游 tag：`v2.7.0`
- 上游 commit：`6de763f30db768970e159b03c86d72e9df53e276`
- 同步时间：2026-09-03（手工同步，CLI 包装层同步补齐 v2.7 新能力）

> **关于 v2.7.1**：上游工作区 `APP_VERSION` 已是 `2.7.1`，但那两个提交（`d5f8071`
> 界面品牌化精修、`fe6ed48` 大字号即时生效）**只改 `gui/app.py` / `version.txt`，
> `core/` 零改动**。按本节「跟已发布 tag、不跟 main 半成品」的规则，本次同步取
> `v2.7.0` tag 的内容——它与当前上游 `core/` 逐字节一致，已是最新内核。
<!-- SYNC:END -->

## 首次种子（2026-08-05，人工完成）

自动同步机制建仓时还没跑过，这份 vendor 代码是**手工从 excel-router 本地工作区**拷贝的种子，
比上游最新发布 tag `v2.5.1`（commit `6d83c1a`）更新：本地工作区当时已经写好但**还没提交/打
tag** 的两项功能也包含在内——

- `core/pdf_dist.py`（PDF 按网格加密分发，全新文件）
- `core/splitter.py` 里的 `keep_formulas` 保留公式支持（在 v2.5.1 基础上的修改）

`.upstream-tag` 仍然记成 `v2.5.1`（最后一个真实发布的 tag），**不是**这份种子代码对应的版本——
这是有意为之：等 AbeLin 把这两项功能提交并打出下一个 tag（比如 `v2.6.0`）后，自动同步会看到
`v2.6.0 > v2.5.1` 才会真正跑一次拷贝+测试+提交，届时 vendor 代码会被替换成那个 tag 的正式内容
（预期与这份种子基本一致，因为种子本来就是从同一份工作区拷的）。**在此之前，自动同步看到远端
仍是 `v2.5.1`（跟记录的一样新），会直接跳过，不会把这份种子回退掉。**

如果你在 `v2.6.0` 真正发布前看到这个文件，说明同步还没发生过一次，属于预期状态。
