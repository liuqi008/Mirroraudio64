比较完美的版本，但windows自启和日志没有实现，输出延迟也略高。
在3.0.6.1的基础上降低输出延迟
独占缓冲对齐策略——按同步方式分档（已完成）
调整位置：BufAligned(...)
函数签名改为：static int BufAligned(int wantMs, bool exclusive, bool useEvent, double defMs, double minMs, BufferAlignMode mode)
核心逻辑：
独占 + 事件（Event/AUTO）：
最小对齐（MinAlign）≥ 1× 最小周期
默认对齐（DefaultAlign）≥ 2× 默认周期
独占 + 轮询（Polling）：
最小对齐 ≥ 2× 最小周期
默认对齐 ≥ 3× 默认周期
共享：默认对齐 ≥ 2× 默认周期
所有调用点均已加入 useEvent 参数：
主通道独占：(_cfg.MainSync == Event || Auto)
主通道共享：传同样参数但分支忽略
副通道独占/共享同理
效果：若设备最小周期为 3 ms，选择“事件+最小对齐”即可下探到 3 ms；填 6/9 ms 也会严格按周期倍数对齐，不再被无谓抬高。
2) 主通道输入队列降到低延迟档（已完成）
调整位置：_bufMain = new BufferedWaveProvider(inFmt) { ... BufferDuration = ... }
由原来的 Max(MainBufMs*4, 80) 改为：
若预期 独占+事件：Max(MainBufMs*2, 24)
否则仍为 Max(MainBufMs*4, 80)（稳健）
并发一条优化：当预期 独占+事件时，显式
_bufMain.ReadFully = false;
降低零填充导致的排队，进一步缩短端到端延迟。
副通道保持原大队列逻辑，保证采集/推流稳定性。
