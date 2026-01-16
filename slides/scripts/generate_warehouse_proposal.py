#!/usr/bin/env python3
"""
Generate Warehouse System Proposal slides.

Usage:
    uv run python slides/scripts/generate_warehouse_proposal.py
"""

from pptx.enum.text import PP_ALIGN
from generate_pptx import SlideGenerator


class WarehouseProposalGenerator(SlideGenerator):
    """Generate Warehouse System Proposal slides."""

    def create_slide_2_purpose(self):
        """Slide 2: 目的 - Key slide showing benefits for both management and operations."""
        slide = self.add_content_slide("なぜWarehouseシステムが必要か")
        c = self.colors

        # Layout: Two columns (Management | Operations) with center connector
        col_width = 7.0
        gap = 2.5
        total_width = col_width * 2 + gap
        start_left = self.center_left(total_width)
        start_top = self.MARGIN_TOP + 0.1

        # === Left Column: 管理側 ===
        mgmt_left = start_left
        self.add_rounded_box(slide, mgmt_left, start_top, col_width, 0.8, c.dark_navy, "👔 管理側のメリット", 24, c.white)

        mgmt_benefits = [
            ("在庫がリアルタイムで見える", "どこに何がいくつあるか、即座に把握"),
            ("滞留を自動検知", "90日動きがなければアラート"),
            ("どこからでも確認", "会議中・出張先からもアクセス可能"),
        ]

        benefit_height = 1.5
        benefit_gap = 0.25
        mgmt_top = start_top + 1.0

        for i, (title, desc) in enumerate(mgmt_benefits):
            top = mgmt_top + i * (benefit_height + benefit_gap)
            self.add_rounded_box(slide, mgmt_left, top, col_width, benefit_height, c.light_gray, "", 14, c.dark_navy)
            self.add_text_box(slide, mgmt_left + 0.3, top + 0.2, col_width - 0.6, 0.5, title, 18, c.dark_navy, True, PP_ALIGN.LEFT)
            self.add_text_box(slide, mgmt_left + 0.3, top + 0.7, col_width - 0.6, 0.6, desc, 14, c.dark_navy, False, PP_ALIGN.LEFT)

        # === Right Column: 現場側 ===
        ops_left = start_left + col_width + gap
        self.add_rounded_box(slide, ops_left, start_top, col_width, 0.8, c.gold, "🔧 現場側のメリット", 24, c.white)

        ops_benefits = [
            ("スキャン1回で作業完了", "手書き・二重入力がゼロに"),
            ("何をすべきか明確", "システムが作業と場所を指示"),
            ("誰でも同じ品質", "ベテラン依存からの脱却"),
        ]

        ops_top = start_top + 1.0

        for i, (title, desc) in enumerate(ops_benefits):
            top = ops_top + i * (benefit_height + benefit_gap)
            self.add_rounded_box(slide, ops_left, top, col_width, benefit_height, c.light_gray, "", 14, c.dark_navy)
            self.add_text_box(slide, ops_left + 0.3, top + 0.2, col_width - 0.6, 0.5, title, 18, c.dark_navy, True, PP_ALIGN.LEFT)
            self.add_text_box(slide, ops_left + 0.3, top + 0.7, col_width - 0.6, 0.6, desc, 14, c.dark_navy, False, PP_ALIGN.LEFT)

        # === Center connector ===
        center_x = start_left + col_width + gap / 2
        connector_top = start_top + 1.5
        self.add_text_box(slide, center_x - 0.5, connector_top, 1.0, 3.0, "⟷", 48, c.gold, True, PP_ALIGN.CENTER)

        # === Bottom: Key message ===
        msg_top = start_top + 5.8
        msg_box_height = 1.2

        self.add_rounded_box(slide, start_left, msg_top, total_width, msg_box_height, c.dark_navy, "", 20, c.white)
        self.add_text_box(slide, start_left, msg_top + 0.15, total_width, 0.5,
                          "「見えない」から「見える」へ", 28, c.gold, True, PP_ALIGN.CENTER)
        self.add_text_box(slide, start_left, msg_top + 0.65, total_width, 0.5,
                          "見えれば判断できる。判断できれば動かせる。", 20, c.white, False, PP_ALIGN.CENTER)

        return slide

    def create_slide_3_issues(self):
        """Slide 3: 現状の課題."""
        slide = self.add_content_slide("現状の課題")
        c = self.colors

        issues = [
            ("見えない", "在庫状況が不明"),
            ("溜まる", "倉庫が満杯に"),
            ("判断できない", "データがない"),
            ("属人的", "担当者依存"),
            ("現場負荷", "手書き・二重入力"),
        ]

        box_width = 3.0
        box_height = 2.8
        gap = 0.3

        total_width = box_width * 5 + gap * 4
        start_left = self.center_left(total_width)
        start_top = self.MARGIN_TOP + 0.3

        for i, (title, desc) in enumerate(issues):
            left = start_left + i * (box_width + gap)
            self.add_text_box(slide, left, start_top, box_width, 0.5, "❌", 32, c.burgundy, True, PP_ALIGN.CENTER)
            self.add_multiline_box(slide, left, start_top + 0.5, box_width, box_height, c.dark_navy, title, desc,
                                   title_size=22, subtitle_size=14)

        return slide

    def create_slide_4_solution(self):
        """Slide 4: 解決の方向性."""
        slide = self.add_content_slide("解決の方向性")
        c = self.colors

        comparisons = [
            ("紙・スプレッドシート管理", "スキャン1回で自動記録"),
            ("倉庫に聞かないとわからない", "どこからでもリアルタイム把握"),
            ("滞留在庫に気づかない", "90日で自動アラート"),
            ("出荷指示がバラバラ", "システムで一元管理"),
            ("紙を見て探し回る", "システムが場所を指示"),
        ]

        col_width = 6.5
        row_height = 1.0
        gap = 0.2
        arrow_width = 0.8

        total_width = col_width * 2 + arrow_width
        start_left = self.center_left(total_width)
        start_top = self.MARGIN_TOP + 0.2

        self.add_rounded_box(slide, start_left, start_top, col_width, 0.7, c.dark_navy, "Before（現状）", 20, c.white)
        self.add_rounded_box(slide, start_left + col_width + arrow_width, start_top, col_width, 0.7, c.gold, "After（システム導入後）", 20, c.white)

        for i, (before, after) in enumerate(comparisons):
            top = start_top + 0.85 + i * (row_height + gap)
            self.add_rounded_box(slide, start_left, top, col_width, row_height, c.light_gray, before, 16, c.dark_navy)
            self.add_text_box(slide, start_left + col_width, top, arrow_width, row_height, "→", 28, c.gold, True, PP_ALIGN.CENTER)
            self.add_rounded_box(slide, start_left + col_width + arrow_width, top, col_width, row_height, c.beige, after, 16, c.dark_navy)

        key_top = start_top + 7.0
        self.add_text_box(slide, start_left, key_top, total_width, 0.5,
                          "全ての入出荷を記録し、止まっている作業を「残」として可視化する。",
                          22, c.gold, True, PP_ALIGN.CENTER)

        return slide

    def create_slide_5_system(self):
        """Slide 5: システム構成イメージ - Management vs Operations view."""
        slide = self.add_content_slide("システム構成イメージ")
        c = self.colors

        # Layout: Two columns with shared system in center
        col_width = 7.0
        gap = 2.5
        total_width = col_width * 2 + gap
        start_left = self.center_left(total_width)
        start_top = self.MARGIN_TOP + 0.1

        # === Left Column: 管理側 (Dashboard) ===
        mgmt_left = start_left
        self.add_rounded_box(slide, mgmt_left, start_top, col_width, 0.8, c.dark_navy, "👔 管理側：ダッシュボード", 22, c.white)

        mgmt_features = [
            ("📊 全体進捗の把握", "入荷・出荷・在庫状況を一覧"),
            ("🔔 アラート通知", "SLA超過・滞留在庫を自動検知"),
            ("📈 データ分析", "滞留傾向・作業効率をレポート"),
            ("✅ 判断・承認", "廃棄/売却の意思決定"),
        ]

        feature_height = 1.2
        feature_gap = 0.2
        mgmt_top = start_top + 1.0

        for i, (title, desc) in enumerate(mgmt_features):
            top = mgmt_top + i * (feature_height + feature_gap)
            self.add_rounded_box(slide, mgmt_left, top, col_width, feature_height, c.light_gray, "", 14, c.dark_navy)
            self.add_text_box(slide, mgmt_left + 0.3, top + 0.15, col_width - 0.6, 0.5, title, 16, c.dark_navy, True, PP_ALIGN.LEFT)
            self.add_text_box(slide, mgmt_left + 0.3, top + 0.6, col_width - 0.6, 0.5, desc, 13, c.dark_navy, False, PP_ALIGN.LEFT)

        # === Right Column: 現場側 (Mobile) ===
        ops_left = start_left + col_width + gap
        self.add_rounded_box(slide, ops_left, start_top, col_width, 0.8, c.gold, "🔧 現場側：モバイルアプリ", 22, c.white)

        ops_features = [
            ("📋 今日の作業一覧", "やるべきタスクが自動表示"),
            ("📍 場所ナビ", "棚番号・ロケーションを指示"),
            ("📷 スキャン完了", "バーコード読取で作業記録"),
            ("✔️ 進捗自動更新", "完了したらリアルタイム反映"),
        ]

        ops_top = start_top + 1.0

        for i, (title, desc) in enumerate(ops_features):
            top = ops_top + i * (feature_height + feature_gap)
            self.add_rounded_box(slide, ops_left, top, col_width, feature_height, c.light_gray, "", 14, c.dark_navy)
            self.add_text_box(slide, ops_left + 0.3, top + 0.15, col_width - 0.6, 0.5, title, 16, c.dark_navy, True, PP_ALIGN.LEFT)
            self.add_text_box(slide, ops_left + 0.3, top + 0.6, col_width - 0.6, 0.5, desc, 13, c.dark_navy, False, PP_ALIGN.LEFT)

        # === Center connector with shared functions ===
        center_x = start_left + col_width
        center_width = gap
        center_top = start_top + 1.5

        # Arrows
        self.add_text_box(slide, center_x, center_top + 0.5, center_width, 0.5, "←→", 28, c.gold, True, PP_ALIGN.CENTER)
        self.add_text_box(slide, center_x, center_top + 1.5, center_width, 0.8, "データ\n連携", 14, c.dark_navy, True, PP_ALIGN.CENTER)
        self.add_text_box(slide, center_x, center_top + 2.5, center_width, 0.5, "←→", 28, c.gold, True, PP_ALIGN.CENTER)

        # === Bottom: Core system functions ===
        bottom_top = start_top + 6.0
        func_width = 5.0
        func_gap = 0.5
        total_func_width = func_width * 3 + func_gap * 2
        func_start_left = self.center_left(total_func_width)

        self.add_text_box(slide, func_start_left, bottom_top - 0.5, total_func_width, 0.4,
                          "共通基盤：3つの管理機能", 16, c.dark_navy, True, PP_ALIGN.CENTER)

        functions = [
            ("📥 入荷管理", "発注→検品→棚入れ"),
            ("📦 在庫管理", "保管・滞留検知"),
            ("📤 出荷管理", "指示→発送→配送"),
        ]

        for i, (title, desc) in enumerate(functions):
            left = func_start_left + i * (func_width + func_gap)
            self.add_rounded_box(slide, left, bottom_top, func_width, 1.0, c.dark_navy, f"{title}", 14, c.white)

        return slide

    def create_slide_6_efficiency(self):
        """Slide 6: 現場の作業効率化."""
        slide = self.add_content_slide("現場の作業効率化")
        c = self.colors

        self.add_text_box(slide, self.content_left, self.MARGIN_TOP - 0.3, self.content_width, 0.4,
                          "スキャン1回で完了、手書き不要", 22, c.dark_navy, False, PP_ALIGN.CENTER)

        comparisons = [
            ("紙で商品を探す", "システムが場所を指示"),
            ("手書き→PC入力", "スキャン1回で完了"),
            ("進捗確認が必要", "リアルタイム共有"),
            ("ベテラン依存", "誰でも同品質"),
        ]

        col_width = 6.5
        row_height = 1.0
        gap = 0.2
        arrow_width = 0.8

        total_width = col_width * 2 + arrow_width
        start_left = self.center_left(total_width)
        start_top = self.MARGIN_TOP + 0.5

        self.add_rounded_box(slide, start_left, start_top, col_width, 0.6, c.dark_navy, "Before", 18, c.white)
        self.add_rounded_box(slide, start_left + col_width + arrow_width, start_top, col_width, 0.6, c.gold, "After", 18, c.white)

        for i, (before, after) in enumerate(comparisons):
            top = start_top + 0.75 + i * (row_height + gap)
            self.add_rounded_box(slide, start_left, top, col_width, row_height, c.light_gray, before, 16, c.dark_navy)
            self.add_text_box(slide, start_left + col_width, top, arrow_width, row_height, "→", 28, c.gold, True, PP_ALIGN.CENTER)
            self.add_rounded_box(slide, start_left + col_width + arrow_width, top, col_width, row_height, c.beige, after, 16, c.dark_navy)

        testimonials = ["探す時間が減った", "迷わない", "記録の手間ゼロ"]
        test_top = start_top + 5.5
        test_width = 4.5
        test_gap = 0.4
        total_test_width = test_width * 3 + test_gap * 2
        test_start_left = self.center_left(total_test_width)

        self.add_text_box(slide, test_start_left, test_top - 0.5, total_test_width, 0.4, "現場の声（想定）:", 18, c.dark_navy, True, PP_ALIGN.LEFT)

        for i, text in enumerate(testimonials):
            left = test_start_left + i * (test_width + test_gap)
            self.add_rounded_box(slide, left, test_top, test_width, 0.7, c.gold, text, 16, c.white)

        return slide

    def create_slide_7_dashboard(self):
        """Slide 7: どこからでも状況確認."""
        slide = self.add_content_slide("どこからでも状況確認")
        c = self.colors

        self.add_text_box(slide, self.content_left, self.MARGIN_TOP - 0.3, self.content_width, 0.4,
                          "どこからでもリアルタイムで把握", 22, c.dark_navy, False, PP_ALIGN.CENTER)

        dash_width = 9.0
        dash_height = 5.0
        dash_left = self.content_left + 0.5
        dash_top = self.MARGIN_TOP + 0.5

        self.add_rounded_box(slide, dash_left, dash_top, dash_width, dash_height, c.light_gray, "", 16, c.dark_navy)
        self.add_text_box(slide, dash_left + 0.3, dash_top + 0.2, dash_width - 0.6, 0.5, "ダッシュボード", 18, c.dark_navy, True, PP_ALIGN.LEFT)

        items = [
            ("入荷", "着荷待ち 23 → 検品中 12", c.dark_navy),
            ("出荷", "準備中 22 → 発送待ち 11", c.dark_navy),
            ("在庫", "良品 4,521 / 滞留 156", c.dark_navy),
            ("アラート", "SLA超過 4 / 滞留 156", c.burgundy),
        ]

        item_top = dash_top + 0.8
        for i, (label, value, color) in enumerate(items):
            top = item_top + i * 1.0
            self.add_rounded_box(slide, dash_left + 0.3, top, 1.8, 0.8, color, label, 14, c.white)
            self.add_text_box(slide, dash_left + 2.3, top + 0.2, 6.5, 0.6, value, 16, c.dark_navy, False, PP_ALIGN.LEFT)

        use_left = dash_left + dash_width + 0.8
        use_width = 6.5

        uses = [
            ("会議中に在庫確認", c.gold),
            ("出張先から出荷確認", c.gold),
            ("朝イチでアラート確認", c.gold),
        ]

        self.add_text_box(slide, use_left, dash_top, use_width, 0.5, "いつでも確認できる:", 18, c.dark_navy, True, PP_ALIGN.LEFT)

        for i, (text, color) in enumerate(uses):
            top = dash_top + 0.6 + i * 1.2
            self.add_rounded_box(slide, use_left, top, use_width, 1.0, color, text, 16, c.white)

        return slide

    def create_slide_8_stagnant(self):
        """Slide 8: 滞留在庫の解消."""
        slide = self.add_content_slide("滞留在庫の解消")
        c = self.colors

        self.add_text_box(slide, self.content_left, self.MARGIN_TOP - 0.3, self.content_width, 0.4,
                          "判断を先送りにできない仕組み", 22, c.dark_navy, False, PP_ALIGN.CENTER)

        flow_width = 15.0
        start_left = self.center_left(flow_width)
        start_top = self.MARGIN_TOP + 0.5

        self.add_text_box(slide, start_left, start_top, flow_width, 0.4, "現状の問題:", 18, c.burgundy, True, PP_ALIGN.LEFT)

        problem_items = ["在庫", "放置", "大量滞留", "倉庫パンク"]
        item_width = 3.4

        for i, text in enumerate(problem_items):
            left = start_left + i * (item_width + 0.4)
            color = c.dark_navy if i < 3 else c.burgundy
            self.add_rounded_box(slide, left, start_top + 0.5, item_width, 1.2, color, text, 18, c.white)
            if i < len(problem_items) - 1:
                self.add_text_box(slide, left + item_width, start_top + 0.85, 0.4, 0.5, "→", 24, c.dark_navy, True, PP_ALIGN.CENTER)

        sol_top = start_top + 2.2
        self.add_text_box(slide, start_left, sol_top, flow_width, 0.4, "システム導入後:", 18, c.gold, True, PP_ALIGN.LEFT)

        solution_items = [
            ("在庫", c.dark_navy),
            ("90日動きなし", c.dark_navy),
            ("自動フラグ", c.dark_navy),
            ("本部に通知", c.gold),
            ("3営業日で判断", c.gold),
            ("実行", c.gold),
        ]

        sol_item_width = 2.4
        for i, (text, color) in enumerate(solution_items):
            left = start_left + i * (sol_item_width + 0.25)
            self.add_rounded_box(slide, left, sol_top + 0.5, sol_item_width, 1.4, color, text, 14, c.white)
            if i < len(solution_items) - 1:
                self.add_text_box(slide, left + sol_item_width, sol_top + 0.95, 0.25, 0.5, "→", 18, c.dark_navy, True, PP_ALIGN.CENTER)

        key_top = sol_top + 2.3
        self.add_text_box(slide, start_left, key_top, flow_width, 0.5,
                          "期限付きの「残」として管理することで、滞留を強制的に解消。", 20, c.gold, True, PP_ALIGN.CENTER)

        return slide

    def create_slide_9_effect(self):
        """Slide 9: 導入効果（定量）."""
        slide = self.add_content_slide("導入効果（定量）")
        c = self.colors

        # Left side: 業務効率の改善
        left_start = self.content_left + 0.3
        left_width = 8.0
        start_top = self.MARGIN_TOP + 0.1

        self.add_text_box(slide, left_start, start_top, left_width, 0.4, "業務効率の改善", 18, c.dark_navy, True, PP_ALIGN.LEFT)

        efficiency_data = [
            ("指標", "現状", "導入後", "改善幅"),
            ("在庫精度", "xx%", "99%以上", "+xx%"),
            ("在庫確認時間", "xx分/回", "即時(<1分)", "-xx%"),
            ("入出荷作業時間", "xx分/件", "xx分/件", "-30%想定"),
            ("記録・入力作業", "xx時間/日", "ほぼゼロ", "-90%想定"),
            ("問い合わせ対応", "xx件/日", "xx件/日", "-50%想定"),
        ]

        eff_col_widths = [2.0, 1.6, 1.8, 1.4]
        eff_row_height = 0.55
        eff_gap = 0.08
        eff_top = start_top + 0.5

        for row_idx, row in enumerate(efficiency_data):
            top = eff_top + row_idx * (eff_row_height + eff_gap)
            col_left = left_start
            for col_idx, cell in enumerate(row):
                if row_idx == 0:
                    color = c.dark_navy if col_idx < 2 else c.gold
                    font_color = c.white
                else:
                    color = c.light_gray
                    font_color = c.dark_navy
                self.add_rounded_box(slide, col_left, top, eff_col_widths[col_idx], eff_row_height, color, cell, 12, font_color)
                col_left += eff_col_widths[col_idx] + eff_gap

        # Right side: コストインパクト
        right_start = left_start + left_width + 0.5
        right_width = 7.5

        self.add_text_box(slide, right_start, start_top, right_width, 0.4, "コストインパクト", 18, c.dark_navy, True, PP_ALIGN.LEFT)

        cost_data = [
            ("項目", "現状(年)", "導入後(年)", "削減効果"),
            ("滞留在庫金額", "$xx万", "$xx万", "$xx万削減"),
            ("廃棄ロス", "$xx万", "$xx万", "$xx万削減"),
            ("人件費(記録)", "$xx万", "$xx万", "$xx万削減"),
            ("合計削減効果", "-", "-", "$xx万/年"),
        ]

        cost_col_widths = [1.8, 1.5, 1.5, 1.6]
        cost_row_height = 0.55
        cost_gap = 0.08
        cost_top = start_top + 0.5

        for row_idx, row in enumerate(cost_data):
            top = cost_top + row_idx * (cost_row_height + cost_gap)
            col_left = right_start
            for col_idx, cell in enumerate(row):
                if row_idx == 0:
                    color = c.dark_navy if col_idx < 2 else c.gold
                    font_color = c.white
                elif row_idx == len(cost_data) - 1:
                    color = c.gold if col_idx == 3 else c.light_gray
                    font_color = c.white if col_idx == 3 else c.dark_navy
                else:
                    color = c.light_gray
                    font_color = c.dark_navy
                self.add_rounded_box(slide, col_left, top, cost_col_widths[col_idx], cost_row_height, color, cell, 12, font_color)
                col_left += cost_col_widths[col_idx] + cost_gap

        # Bottom: ROI試算
        roi_top = start_top + 4.2
        roi_width = 16.0
        roi_left = self.center_left(roi_width)

        self.add_text_box(slide, roi_left, roi_top, roi_width, 0.4, "ROI試算", 18, c.dark_navy, True, PP_ALIGN.LEFT)

        roi_items = [
            ("初期投資", "$xx万", c.dark_navy),
            ("年間運用コスト", "$xx万", c.dark_navy),
            ("年間削減効果", "$xx万", c.gold),
            ("投資回収期間", "xx年", c.gold),
        ]

        roi_item_width = 3.6
        roi_gap = 0.4
        roi_box_top = roi_top + 0.5

        for i, (label, value, color) in enumerate(roi_items):
            left = roi_left + i * (roi_item_width + roi_gap)
            self.add_rounded_box(slide, left, roi_box_top, roi_item_width, 0.9, color, f"{label}\n{value}", 14, c.white)

        # Goal message
        goal_top = roi_box_top + 1.2
        goal = "最終ゴール: Pull型（現場任せ）からPush型（本部主導）へ"
        self.add_text_box(slide, roi_left, goal_top, roi_width, 0.5, goal, 18, c.gold, True, PP_ALIGN.CENTER)

        return slide

    def create_slide_10_summary(self):
        """Slide 10: まとめとNext Steps."""
        slide = self.add_content_slide("まとめとNext Steps")
        c = self.colors

        col_width = 5.0
        gap = 0.4
        total_width = col_width * 3 + gap * 2
        start_left = self.center_left(total_width)
        start_top = self.MARGIN_TOP + 0.1

        self.add_rounded_box(slide, start_left, start_top, col_width, 0.6, c.dark_navy, "1. 課題", 20, c.white)
        issues = "・在庫が見えない\n・溜まる\n・判断できない\n・属人的\n・現場負荷が高い"
        self.add_rounded_box(slide, start_left, start_top + 0.7, col_width, 2.8, c.light_gray, issues, 16, c.dark_navy)

        col2_left = start_left + col_width + gap
        self.add_rounded_box(slide, col2_left, start_top, col_width, 0.6, c.dark_navy, "2. 解決策", 20, c.white)
        solutions = "・入出荷をシステム記録\n・「残」として可視化\n・スキャン1回で完了"
        self.add_rounded_box(slide, col2_left, start_top + 0.7, col_width, 2.8, c.light_gray, solutions, 16, c.dark_navy)

        col3_left = col2_left + col_width + gap
        self.add_rounded_box(slide, col3_left, start_top, col_width, 0.6, c.gold, "3. 期待効果", 20, c.white)
        effects = "・在庫リアルタイム把握\n・現場作業の効率化\n・滞留の自動検知\n・Push型オペレーション"
        self.add_rounded_box(slide, col3_left, start_top + 0.7, col_width, 2.8, c.light_gray, effects, 16, c.dark_navy)

        next_top = start_top + 4.0
        self.add_text_box(slide, start_left, next_top, total_width, 0.5, "Next Steps:", 22, c.dark_navy, True, PP_ALIGN.LEFT)

        steps = [("1", "本提案の方向性承認"), ("2", "詳細設計（画面・データ項目）"), ("3", "Phase 1 開発着手")]
        step_width = 5.0
        step_gap = 0.3
        step_top = next_top + 0.6

        for i, (num, text) in enumerate(steps):
            left = start_left + i * (step_width + step_gap)
            self.add_rounded_box(slide, left, step_top, 0.6, 0.6, c.gold, num, 18, c.white)
            self.add_text_box(slide, left + 0.7, step_top + 0.1, step_width - 0.8, 0.5, text, 18, c.dark_navy, False, PP_ALIGN.LEFT)

        return slide

    def generate_all(self):
        """Generate all slides."""
        self.load_template()
        self.delete_all_slides()
        print("Deleted existing slides from template")

        self.add_title_slide(
            "Warehouseシステム構築提案",
            "倉庫業務の可視化による滞留在庫の解消と\nPush型オペレーションの実現"
        )
        print("Created slide 1: Title")

        self.create_slide_2_purpose()
        print("Created slide 2: 目的")

        self.create_slide_3_issues()
        print("Created slide 3: 現状の課題")

        self.create_slide_4_solution()
        print("Created slide 4: 解決の方向性")

        self.create_slide_5_system()
        print("Created slide 5: システム構成イメージ")

        self.create_slide_6_efficiency()
        print("Created slide 6: 現場の作業効率化")

        self.create_slide_7_dashboard()
        print("Created slide 7: どこからでも状況確認")

        self.create_slide_8_stagnant()
        print("Created slide 8: 滞留在庫の解消")

        self.create_slide_9_effect()
        print("Created slide 9: 導入効果")

        self.create_slide_10_summary()
        print("Created slide 10: まとめとNext Steps")


def main():
    template_path = './slides/templates/genda.pptx'
    output_path = './slides/output/2026-01-16_warehouse-system-proposal.pptx'

    gen = WarehouseProposalGenerator(template_path)
    gen.generate_all()
    gen.save(output_path)

    print(f"\nSaved to: {output_path}")
    print(f"Total slides: {len(gen.prs.slides)}")
    print(f"Content area: {gen.content_left}in - {gen.content_right}in (width: {gen.content_width}in)")


if __name__ == "__main__":
    main()
