# 四类零件计算输入

生成报价 JSON 前读取本文件。所有数值使用可计算数字，不要把单位拼入数值。无法确认的关键输入要保守估算并加入人工复核事项。

## 目录

- 通用字段
- 加工件
- 钣金件
- 焊接件
- 型材组装件
- 分类与默认值审核

## 通用字段

每个 `parts[]` 项目至少包含：

- `part_type`：只允许 `加工件`、`钣金件`、`焊接件`、`型材组装件`。
- `source_file`：原始图纸路径。
- `preview_image`：可选的 PNG/JPG 预览路径。
- `project_code`、`part_no`、`drawing_no`。
- `material`、`quantity`。
- `calculation_inputs`：以下类别专用输入。
- `notes`、`needs_review_fields`、`review_items`。

## 加工件

材料计算：`密度 × 备料长 × 备料宽 × 备料高 × 材料单价`。

输入字段：

- `stock_length`、`stock_width`、`stock_height`：备料尺寸 mm。缺失时可由成品 `length`、`width`、`height` 加余量后推算并复核。
- `unit`：默认 `KG`。
- `process_hours`：按工艺名称填工时或计价驱动量。

`process_hours`支持：`车床`、`铣床`、`磨床`、`钳工`、`其它`、`放电`、`快丝`、`慢丝`、`夹头`、`精雕机`、`CNC`、`镀铬`、`热处理`、`阳极`、`发黑`。

其中车/铣/磨/钳/放电/线切割/精雕/CNC通常为小时，夹头为治具金额，表面处理通常为 kg。不要把所有驱动量都解释为工时。

## 钣金件

材料计算：`展开面积 × 板厚 × 密度 ÷ 材料利用率 × 材料单价`。

`calculation_inputs`字段：

- `unfolded_area_m2`：展开面积 m²，关键必需。
- `sheet_thickness_mm`：板厚 mm，关键必需。
- `material_utilization`：材料利用率 0～1。
- `cut_length_m`：切割长度 m。
- `pierce_count`：穿孔数。
- `cutting_speed_m_per_min`：切割速度 m/min。
- `pierce_seconds`：单次穿孔秒数。
- `cutting_setup_hours_per_batch`：每批切割准备工时。
- `bend_count`：每件折弯数。
- `minutes_per_bend`：单刀折弯分钟。
- `bend_setup_hours_per_batch`：每批折弯准备工时。
- `hardware_cost_each`：五金/压铆件单件金额。
- `manual_minutes_each`：去毛刺、压铆、整形等手工分钟/件。
- `surface_treatment_cost_each`：表处费/件。
- `other_cost_each`：其他费/件。

若没有展开面积，不要使用外包络长×宽代替而不标记复核。切割速度、穿孔时间和利用率与材料及设备相关，AI估值必须标记人工确认。

## 焊接件

材料计算：`BOM净重 ÷ 材料利用率 × 材料单价`。焊接工时由焊缝长度、焊速和电弧作业率驱动。

`calculation_inputs`字段：

- `bom_net_weight_kg`：每套 BOM 净重 kg，关键必需。
- `material_utilization`：材料利用率。
- `purchased_parts_cost_per_set`：外购件/套。
- `prep_hours_per_batch`：备料工时/批。
- `fitup_hours_per_set`：组立工时/套。
- `weld_length_m`：焊缝长度 m/套，关键必需。
- `weld_feature_mm`：焊脚或焊缝特征尺寸 mm，关键必需。
- `cross_section_factor`：焊缝截面积系数。
- `weld_speed_m_per_min`：焊速 m/min。
- `arc_duty_cycle`：电弧作业率 0～1。
- `filler_material_unit_price`：焊材单价。
- `deposition_efficiency`：熔敷效率 0～1。
- `gas_rate_per_hour`：气体费率元/h。
- `grinding_straightening_hours`：打磨校形 h/套。
- `surface_outsource_cost_per_set`：表处/外协费/套。
- `inspection_packaging_cost_per_set`：检验包装费/套。

从 BOM 汇总重量，不要用焊接总成外包络体积代替。焊缝类型、坡口、夹具、变形控制和焊后处理不明确时必须复核。

## 型材组装件

材料计算：按标准料长、理论 kg/m、排料利用率和批量向上取整采购支数，再分摊到每套。

`calculation_inputs`字段：

- `net_length_m_per_set`：每套净长度 m，关键必需。
- `weight_kg_per_m`：理论重量 kg/m，关键必需。
- `standard_bar_length_m`：标准料长 m/支。
- `nesting_utilization`：排料利用率 0～1。
- `cuts_per_set`：切断数/套。
- `minutes_per_cut`：单刀分钟。
- `saw_setup_hours_per_batch`：锯切准备 h/批。
- `holes_per_set`：钻孔数/套。
- `minutes_per_drilled_hole`：钻孔单孔分钟。
- `taps_per_set`：攻牙数/套。
- `minutes_per_tapped_hole`：攻牙单孔分钟。
- `connector_cost_per_set`：连接件/套。
- `fastener_cost_per_set`：紧固件/套。
- `assembly_minutes_per_set`：装配分钟/套。
- `supplemental_weld_minutes_per_set`：补充焊接分钟/套。
- `surface_outsource_cost_per_set`：表处/外协费/套。
- `inspection_packaging_cost_per_set`：检验包装费/套。

确认余料是否可跨订单回用。若不可回用，当前订单承担向上取整后的采购支数；排料利用率和标准料长必须复核。

## 分类与默认值审核

生成器仅为防止公式报错提供少量兼容默认值，例如钣金切割速度、焊接作业率、标准型材长度。任何使用默认值的关键字段都会写入人工复核提醒。

- 未提供 `part_type`但存在唯一类别专用字段时，可以暂时推断类别并标记 `待复核`。
- 无法推断或类别值无效时，兼容模式暂按 `加工件`，并标记 `高风险`。
- 正式报价前必须确认类别和所有高风险输入。分类错误会同时改变材料、工序和批量分摊逻辑，是影响报价偏差和 CPK 的首要因素。
