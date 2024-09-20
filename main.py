import customtkinter

class RiskManagement(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        self.geometry("400x150")
        self.button = customtkinter.CTkButton(self, text= "Accept", command= self.button_callback)
        self.button.pack(padx = 20, pady= 20)

    def button_callback(self):
        print("Button Clicked")
        

app = RiskManagement()
app.mainloop()