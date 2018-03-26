from cx_Freeze import setup, Executable

base = None


executables = [Executable("GetBigorderSmallOrderSent.py", base=base)]

packages = ["idna"]
options = {
    'build_exe': {

        'packages':packages,
    },

}

setup(
    name = "Bai Mengyang",
    options = options,
    version = "1",
    description = 'For Vince',
    executables = executables
)