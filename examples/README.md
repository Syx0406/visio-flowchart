# 示例说明

每个示例都是一个可直接用引擎跑通的流程 JSON：

```bash
cd ..   # 回到 skill 根目录
python generate.py examples/<示例名>.json -o out.vsdx --png
```

| 示例 | 演示的能力 |
|------|----------|
| [leap_year.json](leap_year.json) | 闰年判断 —— 判断分支、是/否标签、V 形分叉汇聚（最经典的流程图结构） |
| [login_check.json](login_check.json) | 用户登录校验 —— 多判断串联、分支汇聚到同一终点、虚线分组框（groups） |

> 想看自动布局（不写 pos）的效果，可把示例里节点的 `pos` 字段删掉再跑，引擎会按层级自动算坐标。
