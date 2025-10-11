# hooks/hook-core.py
from PyInstaller.utils.hooks import collect_submodules, collect_data_files

# 把 core 包的所有子模块都打进去（包含 core.core_voice 等）
hiddenimports = collect_submodules('core')

# 如果 core 里有数据文件（例如词典、模板、json 等），也一并收集
datas = collect_data_files('core')
