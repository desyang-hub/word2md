"""支持 ``python -m word2md`` 运行。"""

import sys

from .cli import main

if __name__ == "__main__":
    sys.exit(main())
