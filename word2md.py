#!/usr/bin/env python3
"""兼容入口：等价于 ``word2md`` 命令，可继续使用 ``python word2md.py xxx.docx``。"""

import sys

from word2md.cli import main

if __name__ == '__main__':
    sys.exit(main())
