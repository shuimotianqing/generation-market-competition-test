import tkinter as tk
import pandas as pd

class SingleChoiceQuizApp:
    def __init__(self, root, file_path):
        self.root = root
        self.root.title("单选题测试")
        
        # 设置窗口大小为固定大小，并拉大高度
        self.root.geometry("400x500")  # 增大了窗口的高度
        
        # 读取 Excel 题库（选择名为“单选题”的工作簿）
        self.df = pd.read_excel(file_path, sheet_name="单选题")
        self.questions = self.df['问题'].tolist()
        
        # 处理选项和答案
        self.options = [self.df[f'选项{i}'].tolist() for i in range(1, 5)]  # 选项1到选项4
        self.answers = self.df['答案'].apply(lambda x: str(x).strip()).tolist()  # 确保答案是字符串
        
        self.current_question = 0
        self.score = 0
        self.user_answers = []
        
        # 题号显示
        self.question_number_label = tk.Label(root, text=f"第 {self.current_question + 1} 题 / 共 {len(self.questions)} 题", font=("Arial", 12))
        self.question_number_label.pack(pady=5)
        
        # 题目显示
        self.question_label = tk.Label(root, text=self.questions[self.current_question], font=("Arial", 14), wraplength=350)  # 字体稍微小一点
        self.question_label.pack(pady=20)
        
        # 选项
        self.var = tk.StringVar(value="None")  # 初始未选中
        self.option_buttons = []
        for i in range(4):
            option_button = tk.Radiobutton(root, text=f"选项{i+1}: {self.options[i][self.current_question]}", variable=self.var, value=str(i+1), command=self.check_answer, anchor="w")
            option_button.pack(pady=5, padx=20, anchor="w")  # 使用pack左对齐，并加入上下间距
            self.option_buttons.append(option_button)
        
        # 结果显示标签
        self.result_label = tk.Label(root, text="", font=("Arial", 12))
        self.result_label.pack(pady=10)
        
        # 输入题号框
        self.jump_label = tk.Label(root, text="输入题号跳转:", font=("Arial", 10))
        self.jump_label.pack(pady=5)

        self.jump_entry = tk.Entry(root, font=("Arial", 12))
        self.jump_entry.pack(pady=5)
        self.jump_entry.bind("<Return>", self.jump_to_question)
        
        # 按钮固定在界面最底部
        self.buttons_frame = tk.Frame(root)
        self.buttons_frame.pack(side=tk.BOTTOM, pady=10, fill="x")
        
        self.prev_button = tk.Button(self.buttons_frame, text="上一题", command=self.prev_question)
        self.prev_button.pack(side=tk.LEFT, padx=20)
        
        self.next_button = tk.Button(self.buttons_frame, text="下一题", command=self.next_question)
        self.next_button.pack(side=tk.RIGHT, padx=20)
    
    def check_answer(self):
        if self.var.get() == "None":
            return
        
        user_answer = self.var.get().strip()
        correct_answer = self.answers[self.current_question]  # 现在保证是 "1", "2", "3", "4"
        
        is_correct = user_answer == correct_answer
        
        if len(self.user_answers) <= self.current_question:
            self.user_answers.append([self.questions[self.current_question], user_answer, correct_answer, is_correct])
        else:
            self.user_answers[self.current_question] = [self.questions[self.current_question], user_answer, correct_answer, is_correct]
        
        # 显示正确或错误
        if is_correct:
            self.result_label.config(text="回答正确！", fg="green")
        else:
            self.result_label.config(text="回答错误！", fg="red")
        
        # 等待0.4秒后跳转到下一题
        self.root.after(400, self.next_question)
    
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
        self.var.set(self.user_answers[self.current_question][1] if len(self.user_answers) > self.current_question else "None")  # 保留之前选择
        is_correct = self.user_answers[self.current_question][3] if len(self.user_answers) > self.current_question else None
        if is_correct is not None:
            self.result_label.config(text="回答正确！" if is_correct else "回答错误！", fg="green" if is_correct else "red")
        else:
            self.result_label.config(text="")
        
        # 更新题号显示
        self.question_number_label.config(text=f"第 {self.current_question + 1} 题 / 共 {len(self.questions)} 题")
        
        # 更新选项文本
        for i, option_button in enumerate(self.option_buttons):
            option_button.config(text=f"选项{i+1}: {self.options[i][self.current_question]}")
    
    def jump_to_question(self, event):
        try:
            target_question = int(self.jump_entry.get()) - 1  # 转换为索引，用户输入的是1到N
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
        
        # 保存结果到 Excel
        result_df = pd.DataFrame(self.user_answers, columns=['问题', '用户答案', '正确答案', '是否正确'])
        result_df.to_excel("答题记录.xlsx", index=False)
        
if __name__ == "__main__":
    file_path = "理论试题1200.xlsx"  # 题库文件名，确保该 Excel 文件存在
    root = tk.Tk()
    app = SingleChoiceQuizApp(root, file_path)
    root.mainloop()
