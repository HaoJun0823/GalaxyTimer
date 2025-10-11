# hooks/hook-core.py
from PyInstaller.utils.hooks import collect_submodules, collect_data_files

# 把本仓库的 core 包及所有子模块/数据都收集进可执行文件
hiddenimports = collect_submodules('core')
datas = collect_data_files('core')
