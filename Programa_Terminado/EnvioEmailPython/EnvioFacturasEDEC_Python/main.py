import customtkinter as ctk

from views.login_view import LoginView


def main():
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")

    app = LoginView()
    app.mainloop()


if __name__ == "__main__":
    main()