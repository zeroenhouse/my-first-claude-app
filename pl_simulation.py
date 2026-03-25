import pandas as pd
import matplotlib.pyplot as plt
import japanize_matplotlib

# PLシミュレーションクラス
class PLSimulation:
    def __init__(self):
        # 基本パラメータ
        self.scenarios = {
            '保守': {'monthly_sales': 20700000, 'annual_sales': 248400000},
            '中央値': {'monthly_sales': 32200000, 'annual_sales': 386400000},
            'モデル店水準': {'monthly_sales': 46000000, 'annual_sales': 552000000}
        }

        # コスト構成（年間）
        self.cost_structure = {
            '景品原価': 0.45,  # 売上の45%
            '人件費': 90000000,  # 月750万円 × 12
            '家賃': 24000000,   # 月200万円 × 12
            '光熱費': 9600000,  # 月80万円 × 12
            '機器保守': 15000000,  # 月125万円 × 12
            'その他雑費': 6000000,  # 月50万円 × 12
            '減価償却費': 20000000  # 機器5年償却
        }

    def calculate_pl(self, scenario_name):
        """指定シナリオのPLを計算"""
        scenario = self.scenarios[scenario_name]
        annual_sales = scenario['annual_sales']

        # コスト計算
        costs = {}
        costs['景品原価'] = annual_sales * self.cost_structure['景品原価']

        # 固定費
        fixed_costs = {k: v for k, v in self.cost_structure.items()
                      if k not in ['景品原価'] and not isinstance(v, float)}
        total_fixed_costs = sum(fixed_costs.values())

        # 営業利益
        operating_profit = annual_sales - costs['景品原価'] - total_fixed_costs

        # 税引前利益
        profit_before_tax = operating_profit - self.cost_structure['減価償却費']

        # 法人税等（35%）
        tax = profit_before_tax * 0.35

        # 当期純利益
        net_profit = profit_before_tax - tax

        return {
            '売上高': annual_sales,
            '景品原価': costs['景品原価'],
            '人件費': self.cost_structure['人件費'],
            '家賃': self.cost_structure['家賃'],
            '光熱費': self.cost_structure['光熱費'],
            '機器保守': self.cost_structure['機器保守'],
            'その他雑費': self.cost_structure['その他雑費'],
            '営業利益': operating_profit,
            '減価償却費': self.cost_structure['減価償却費'],
            '税引前利益': profit_before_tax,
            '法人税等': tax,
            '当期純利益': net_profit
        }

    def create_pl_table(self):
        """全シナリオのPL比較表を作成"""
        results = {}
        for scenario in self.scenarios.keys():
            results[scenario] = self.calculate_pl(scenario)

        # DataFrame作成
        df = pd.DataFrame(results).T
        df = df.map(lambda x: f"{x:,.0f}円" if isinstance(x, (int, float)) else x)

        return df

    def create_profit_chart(self):
        """純利益の比較チャートを作成"""
        profits = {}
        for scenario in self.scenarios.keys():
            pl = self.calculate_pl(scenario)
            profits[scenario] = pl['当期純利益']

        plt.figure(figsize=(10, 6))
        bars = plt.bar(profits.keys(), profits.values(), color=['#ff9999', '#66b3ff', '#99ff99'])

        plt.title('GiGO我孫子 PLシミュレーション - 当期純利益比較', fontsize=16, pad=20)
        plt.ylabel('当期純利益 (円)', fontsize=12)
        plt.xlabel('シナリオ', fontsize=12)

        # 値ラベル追加
        for bar, value in zip(bars, profits.values()):
            plt.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1000000,
                    f'{value:,.0f}円', ha='center', va='bottom', fontsize=10)

        plt.grid(axis='y', alpha=0.3)
        plt.tight_layout()
        plt.savefig('pl_profit_comparison.png', dpi=300, bbox_inches='tight')
        plt.show()

        print("PL純利益比較チャートを保存しました: pl_profit_comparison.png")

    def create_cost_breakdown_chart(self, scenario_name='中央値'):
        """指定シナリオのコスト内訳円グラフを作成"""
        pl = self.calculate_pl(scenario_name)

        # コスト内訳（売上を除く）
        costs = {
            '景品原価': pl['景品原価'],
            '人件費': pl['人件費'],
            '家賃': pl['家賃'],
            '光熱費': pl['光熱費'],
            '機器保守': pl['機器保守'],
            'その他雑費': pl['その他雑費']
        }

        plt.figure(figsize=(10, 8))
        plt.pie(costs.values(), labels=costs.keys(), autopct='%1.1f%%', startangle=90)
        plt.title(f'GiGO我孫子 コスト内訳 - {scenario_name}シナリオ', fontsize=16, pad=20)
        plt.axis('equal')
        plt.tight_layout()
        plt.savefig(f'pl_cost_breakdown_{scenario_name}.png', dpi=300, bbox_inches='tight')
        plt.show()

        print(f"コスト内訳チャートを保存しました: pl_cost_breakdown_{scenario_name}.png")

# メイン実行
if __name__ == "__main__":
    sim = PLSimulation()

    print("=== GiGO我孫子 PLシミュレーション ===\n")

    # PL比較表表示
    print("PL比較表:")
    pl_table = sim.create_pl_table()
    print(pl_table)
    print("\n")

    # 各シナリオの詳細表示
    for scenario in sim.scenarios.keys():
        print(f"=== {scenario}シナリオ 詳細 ===")
        pl = sim.calculate_pl(scenario)
        for key, value in pl.items():
            if isinstance(value, (int, float)):
                print(f"{key}: {value:,.0f}円")
            else:
                print(f"{key}: {value}")
        print(f"利益率: {pl['当期純利益']/pl['売上高']:.1%}")
        print("\n")

    # チャート生成
    print("チャート生成中...")
    sim.create_profit_chart()
    sim.create_cost_breakdown_chart('中央値')