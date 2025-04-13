import tkinter as tk
import pandas as pd

class MultipleChoiceQuizApp:
    def __init__(self, root, file_path):
        self.root = root
        self.root.title("多选题测试")
        self.root.geometry("600x600")  # 调整窗口大小
        
        # 读取 Excel 题库（选择名为“多选题”的工作簿）
        self.df = pd.read_excel(file_path, sheet_name="多选题")
        self.questions = self.df['问题'].tolist()
        
        # 获取所有选项
        self.options = self.df[['选项1', '选项2', '选项3', '选项4', '选项5']].fillna('').values.tolist()
        
        # 解析正确答案（将"1|3"转换为列表["1", "3"]）
        self.answers = self.df['答案'].astype(str).apply(lambda x: x.split('|')).tolist()
        
        self.current_question = 0
        self.score = 0
        self.user_answers = []

        # 题号显示
        self.question_number_label = tk.Label(root, text=f"第 {self.current_question + 1} 题 / 共 {len(self.questions)} 题", font=("Arial", 12))
        self.question_number_label.pack(pady=5)

        # 题目显示
        self.question_label = tk.Label(root, text=self.questions[self.current_question], font=("Arial", 12), wraplength=550)
        self.question_label.pack(pady=20)

        # 复选框变量
        self.var_list = [tk.BooleanVar(value=False) for _ in range(5)]
        self.checkboxes = [tk.Checkbutton(root, text="", variable=self.var_list[i], anchor="w") for i in range(5)]
        
        for checkbox in self.checkboxes:
            checkbox.pack(pady=5, padx=20, anchor="w")

        # 结果显示
        self.result_label = tk.Label(root, text="", font=("Arial", 12))
        self.result_label.pack(pady=10)

        # 输入题号框
        self.jump_label = tk.Label(root, text="输入题号跳转:", font=("Arial", 10))
        self.jump_label.pack(pady=5)

        self.jump_entry = tk.Entry(root, font=("Arial", 12), width=5, justify="center")
        self.jump_entry.pack(pady=5)
        self.jump_entry.bind("<Return>", self.jump_to_question)

        # 按钮框架（用于对齐）
        self.buttons_frame = tk.Frame(root)
        self.buttons_frame.pack(side=tk.BOTTOM, pady=10, fill="x")

        self.prev_button = tk.Button(self.buttons_frame, text="上一题", command=self.prev_question, width=10)
        self.prev_button.pack(side=tk.LEFT, padx=20)

        self.submit_button = tk.Button(self.buttons_frame, text="确定", command=self.check_answer, width=10)
        self.submit_button.pack(side=tk.LEFT, padx=20)

        self.next_button = tk.Button(self.buttons_frame, text="下一题", command=self.next_question, width=10)
        self.next_button.pack(side=tk.LEFT, padx=20)

        # 加载第一个题目
        self.update_question()

    def check_answer(self):
        selected_answers = [str(i + 1) for i, var in enumerate(self.var_list) if var.get()]
        correct_answers = self.answers[self.current_question]

        is_correct = sorted(selected_answers) == sorted(correct_answers)

        if len(self.user_answers) <= self.current_question:
            self.user_answers.append([self.questions[self.current_question], selected_answers, correct_answers, is_correct])
        else:
            self.user_answers[self.current_question] = [self.questions[self.current_question], selected_answers, correct_answers, is_correct]

        self.result_label.config(text="回答正确！" if is_correct else "回答错误！", fg="green" if is_correct else "red")

    def prev_question(self):
        if self.current_question > 0:
            self.current_question -= 1
            self.update_question()

    def next_question(self):
        if self.current_question < len(self.questions) - 1:
            self.current_question += 1
            self.update_question()
        else:
            self.show_result()

    def update_question(self):
        self.question_label.config(text=self.questions[self.current_question])

        for i in range(5):
            option_text = self.options[self.current_question][i]
            if option_text:
                self.checkboxes[i].config(text=option_text, state="normal")
            else:
                self.checkboxes[i].config(text="", state="disabled")

        self.question_number_label.config(text=f"第 {self.current_question + 1} 题 / 共 {len(self.questions)} 题")

        last_answers = self.user_answers[self.current_question][1] if len(self.user_answers) > self.current_question else []
        for i in range(5):
            self.var_list[i].set(str(i + 1) in last_answers)

        is_correct = self.user_answers[self.current_question][3] if len(self.user_answers) > self.current_question else None
        if is_correct is not None:
            self.result_label.config(text="回答正确！" if is_correct else "回答错误！", fg="green" if is_correct else "red")
        else:
            self.result_label.config(text="")

    def jump_to_question(self, event):
        try:
            target_question = int(self.jump_entry.get()) - 1
            if 0 <= target_question < len(self.questions):
                self.current_question = target_question
                self.update_question()
            else:
                self.result_label.config(text="请输入有效的题号！", fg="red")
        except ValueError:
            self.result_label.config(text="请输入有效的数字！", fg="red")

    def show_result(self):
        correct_count = sum(1 for ans in self.user_answers if ans[3])
        self.result_label.config(text=f"测试完成，你的得分是: {correct_count}/{len(self.questions)}", fg="blue")

        result_df = pd.DataFrame(self.user_answers, columns=['问题', '用户答案', '正确答案', '是否正确'])
        result_df.to_excel("多选题答题记录.xlsx", index=False)

if __name__ == "__main__":
    file_path = "理论试题1200.xlsx"
    root = tk.Tk()
    app = MultipleChoiceQuizApp(root, file_path)
    root.mainloop()
