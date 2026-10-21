"""Recover ordinary Excel formulas from BIFF8 and BIFF12 RPN tokens.

Callers resolve PtgExp using their shared/array formula records and supply only
local workbook references. Unsupported expressions fail instead of inventing a
formula, allowing textconv to retain a hash of the original bytes.
"""

import math
import re
import struct
from dataclasses import dataclass
from decimal import Decimal


class FormulaError(ValueError):
    """A formula is malformed or contains an unsupported BIFF token."""


# MS-XLSB 2.5.98.10 and MS-XLS 2.5.198.17 (name, minimum, maximum arity).
# https://learn.microsoft.com/openspecs/office_file_formats/ms-xlsb/90a52fcb-ce63-497f-a3d3-173c42d82242
FTAB = {
    0: ("COUNT", 1, 255),
    1: ("IF", 2, 3),
    2: ("ISNA", 1, 1),
    3: ("ISERROR", 1, 1),
    4: ("SUM", 1, 255),
    5: ("AVERAGE", 1, 255),
    6: ("MIN", 1, 255),
    7: ("MAX", 1, 255),
    8: ("ROW", 0, 1),
    9: ("COLUMN", 0, 1),
    10: ("NA", 0, 0),
    11: ("NPV", 2, 255),
    12: ("STDEV", 1, 255),
    13: ("DOLLAR", 1, 2),
    14: ("FIXED", 1, 3),
    15: ("SIN", 1, 1),
    16: ("COS", 1, 1),
    17: ("TAN", 1, 1),
    18: ("ATAN", 1, 1),
    19: ("PI", 0, 0),
    20: ("SQRT", 1, 1),
    21: ("EXP", 1, 1),
    22: ("LN", 1, 1),
    23: ("LOG10", 1, 1),
    24: ("ABS", 1, 1),
    25: ("INT", 1, 1),
    26: ("SIGN", 1, 1),
    27: ("ROUND", 2, 2),
    28: ("LOOKUP", 2, 3),
    29: ("INDEX", 2, 4),
    30: ("REPT", 2, 2),
    31: ("MID", 3, 3),
    32: ("LEN", 1, 1),
    33: ("VALUE", 1, 1),
    34: ("TRUE", 0, 0),
    35: ("FALSE", 0, 0),
    36: ("AND", 1, 255),
    37: ("OR", 1, 255),
    38: ("NOT", 1, 1),
    39: ("MOD", 2, 2),
    40: ("DCOUNT", 3, 3),
    41: ("DSUM", 3, 3),
    42: ("DAVERAGE", 3, 3),
    43: ("DMIN", 3, 3),
    44: ("DMAX", 3, 3),
    45: ("DSTDEV", 3, 3),
    46: ("VAR", 1, 255),
    47: ("DVAR", 3, 3),
    48: ("TEXT", 2, 2),
    49: ("LINEST", 1, 4),
    50: ("TREND", 1, 4),
    51: ("LOGEST", 1, 4),
    52: ("GROWTH", 1, 4),
    53: ("GOTO", 1, 1),
    54: ("HALT", 0, 1),
    55: ("RETURN", 0, 1),
    56: ("PV", 3, 5),
    57: ("FV", 3, 5),
    58: ("NPER", 3, 5),
    59: ("PMT", 3, 5),
    60: ("RATE", 3, 6),
    61: ("MIRR", 3, 3),
    62: ("IRR", 1, 2),
    63: ("RAND", 0, 0),
    64: ("MATCH", 2, 3),
    65: ("DATE", 3, 3),
    66: ("TIME", 3, 3),
    67: ("DAY", 1, 1),
    68: ("MONTH", 1, 1),
    69: ("YEAR", 1, 1),
    70: ("WEEKDAY", 1, 2),
    71: ("HOUR", 1, 1),
    72: ("MINUTE", 1, 1),
    73: ("SECOND", 1, 1),
    74: ("NOW", 0, 0),
    75: ("AREAS", 1, 1),
    76: ("ROWS", 1, 1),
    77: ("COLUMNS", 1, 1),
    78: ("OFFSET", 3, 5),
    79: ("ABSREF", 2, 2),
    80: ("RELREF", 2, 2),
    81: ("ARGUMENT", 0, 3),
    82: ("SEARCH", 2, 3),
    83: ("TRANSPOSE", 1, 1),
    84: ("ERROR", 0, 2),
    85: ("STEP", 0, 0),
    86: ("TYPE", 1, 1),
    87: ("ECHO", 0, 1),
    88: ("SET.NAME", 1, 2),
    89: ("CALLER", 0, 0),
    90: ("DEREF", 1, 1),
    91: ("WINDOWS", 0, 2),
    93: ("DOCUMENTS", 0, 2),
    94: ("ACTIVE.CELL", 0, 0),
    95: ("SELECTION", 0, 0),
    96: ("RESULT", 0, 1),
    97: ("ATAN2", 2, 2),
    98: ("ASIN", 1, 1),
    99: ("ACOS", 1, 1),
    100: ("CHOOSE", 2, 255),
    101: ("HLOOKUP", 3, 4),
    102: ("VLOOKUP", 3, 4),
    103: ("LINKS", 0, 2),
    104: ("INPUT", 1, 7),
    105: ("ISREF", 1, 1),
    106: ("GET.FORMULA", 1, 1),
    107: ("GET.NAME", 1, 2),
    108: ("SET.VALUE", 2, 2),
    109: ("LOG", 1, 2),
    110: ("EXEC", 1, 4),
    111: ("CHAR", 1, 1),
    112: ("LOWER", 1, 1),
    113: ("UPPER", 1, 1),
    114: ("PROPER", 1, 1),
    115: ("LEFT", 1, 2),
    116: ("RIGHT", 1, 2),
    117: ("EXACT", 1, 2),
    118: ("TRIM", 1, 1),
    119: ("REPLACE", 4, 4),
    120: ("SUBSTITUTE", 3, 4),
    121: ("CODE", 1, 1),
    122: ("NAMES", 0, 3),
    123: ("DIRECTORY", 0, 1),
    124: ("FIND", 2, 3),
    125: ("CELL", 1, 2),
    126: ("ISERR", 1, 1),
    127: ("ISTEXT", 1, 1),
    128: ("ISNUMBER", 1, 1),
    129: ("ISBLANK", 1, 1),
    130: ("T", 1, 1),
    131: ("N", 1, 1),
    132: ("FOPEN", 1, 2),
    133: ("FCLOSE", 1, 1),
    134: ("FSIZE", 1, 1),
    135: ("FREADLN", 1, 1),
    136: ("FREAD", 2, 2),
    137: ("FWRITELN", 2, 2),
    138: ("FWRITE", 2, 2),
    139: ("FPOS", 1, 2),
    140: ("DATEVALUE", 1, 1),
    141: ("TIMEVALUE", 1, 1),
    142: ("SLN", 3, 3),
    143: ("SYD", 4, 4),
    144: ("DDB", 4, 5),
    145: ("GET.DEF", 1, 3),
    146: ("REFTEXT", 1, 2),
    147: ("TEXTREF", 1, 2),
    148: ("INDIRECT", 1, 2),
    149: ("REGISTER", 1, 255),
    150: ("CALL", 1, 255),
    151: ("ADD.BAR", 0, 1),
    152: ("ADD.MENU", 2, 4),
    153: ("ADD.COMMAND", 3, 5),
    154: ("ENABLE.COMMAND", 4, 5),
    155: ("CHECK.COMMAND", 4, 5),
    156: ("RENAME.COMMAND", 4, 5),
    157: ("SHOW.BAR", 0, 1),
    158: ("DELETE.MENU", 2, 3),
    159: ("DELETE.COMMAND", 3, 4),
    160: ("GET.CHART.ITEM", 1, 3),
    161: ("DIALOG.BOX", 1, 1),
    162: ("CLEAN", 1, 1),
    163: ("MDETERM", 1, 1),
    164: ("MINVERSE", 1, 1),
    165: ("MMULT", 2, 2),
    166: ("FILES", 0, 2),
    167: ("IPMT", 4, 6),
    168: ("PPMT", 4, 6),
    169: ("COUNTA", 1, 255),
    170: ("CANCEL.KEY", 0, 2),
    171: ("FOR", 3, 4),
    172: ("WHILE", 1, 1),
    173: ("BREAK", 0, 0),
    174: ("NEXT", 0, 0),
    175: ("INITIATE", 2, 2),
    176: ("REQUEST", 2, 2),
    177: ("POKE", 3, 3),
    178: ("EXECUTE", 2, 2),
    179: ("TERMINATE", 1, 1),
    180: ("RESTART", 0, 1),
    181: ("HELP", 0, 1),
    182: ("GET.BAR", 0, 4),
    183: ("PRODUCT", 1, 255),
    184: ("FACT", 1, 1),
    185: ("GET.CELL", 1, 2),
    186: ("GET.WORKSPACE", 1, 1),
    187: ("GET.WINDOW", 1, 2),
    188: ("GET.DOCUMENT", 1, 2),
    189: ("DPRODUCT", 3, 3),
    190: ("ISNONTEXT", 1, 1),
    191: ("GET.NOTE", 0, 3),
    192: ("NOTE", 0, 4),
    193: ("STDEVP", 1, 255),
    194: ("VARP", 1, 255),
    195: ("DSTDEVP", 3, 3),
    196: ("DVARP", 3, 3),
    197: ("TRUNC", 1, 2),
    198: ("ISLOGICAL", 1, 1),
    199: ("DCOUNTA", 3, 3),
    200: ("DELETE.BAR", 1, 1),
    201: ("UNREGISTER", 1, 1),
    204: ("USDOLLAR", 1, 2),
    205: ("FINDB", 2, 3),
    206: ("SEARCHB", 2, 3),
    207: ("REPLACEB", 4, 4),
    208: ("LEFTB", 1, 2),
    209: ("RIGHTB", 1, 2),
    210: ("MIDB", 3, 3),
    211: ("LENB", 3, 3),
    212: ("ROUNDUP", 2, 2),
    213: ("ROUNDDOWN", 2, 2),
    214: ("ASC", 1, 1),
    215: ("DBCS", 1, 1),
    216: ("RANK", 2, 3),
    219: ("ADDRESS", 2, 5),
    220: ("DAYS360", 2, 3),
    221: ("TODAY", 0, 0),
    222: ("VDB", 5, 7),
    223: ("ELSE", 0, 0),
    224: ("ELSE.IF", 1, 1),
    225: ("END.IF", 0, 0),
    226: ("FOR.CELL", 1, 3),
    227: ("MEDIAN", 1, 255),
    228: ("SUMPRODUCT", 1, 255),
    229: ("SINH", 1, 1),
    230: ("COSH", 1, 1),
    231: ("TANH", 1, 1),
    232: ("ASINH", 1, 1),
    233: ("ACOSH", 1, 1),
    234: ("ATANH", 1, 1),
    235: ("DGET", 3, 3),
    236: ("CREATE.OBJECT", 2, 11),
    237: ("VOLATILE", 0, 1),
    238: ("LAST.ERROR", 0, 0),
    239: ("CUSTOM.UNDO", 0, 2),
    240: ("CUSTOM.REPEAT", 0, 3),
    241: ("FORMULA.CONVERT", 2, 5),
    242: ("GET.LINK.INFO", 2, 4),
    243: ("TEXT.BOX", 1, 4),
    244: ("INFO", 1, 1),
    245: ("GROUP", 0, 0),
    246: ("GET.OBJECT", 1, 5),
    247: ("DB", 4, 5),
    248: ("PAUSE", 0, 1),
    251: ("RESUME", 0, 1),
    252: ("FREQUENCY", 2, 2),
    253: ("ADD.TOOLBAR", 0, 2),
    254: ("DELETE.TOOLBAR", 1, 1),
    256: ("RESET.TOOLBAR", 1, 1),
    257: ("EVALUATE", 1, 1),
    258: ("GET.TOOLBAR", 1, 2),
    259: ("GET.TOOL", 1, 3),
    260: ("SPELLING.CHECK", 1, 3),
    261: ("ERROR.TYPE", 1, 1),
    262: ("APP.TITLE", 0, 1),
    263: ("WINDOW.TITLE", 0, 1),
    264: ("SAVE.TOOLBAR", 0, 2),
    265: ("ENABLE.TOOL", 3, 3),
    266: ("PRESS.TOOL", 3, 3),
    267: ("REGISTER.ID", 2, 3),
    268: ("GET.WORKBOOK", 1, 2),
    269: ("AVEDEV", 1, 255),
    270: ("BETADIST", 3, 5),
    271: ("GAMMALN", 1, 1),
    272: ("BETAINV", 3, 5),
    273: ("BINOMDIST", 4, 4),
    274: ("CHIDIST", 2, 2),
    275: ("CHIINV", 2, 2),
    276: ("COMBIN", 2, 2),
    277: ("CONFIDENCE", 3, 3),
    278: ("CRITBINOM", 3, 3),
    279: ("EVEN", 1, 1),
    280: ("EXPONDIST", 3, 3),
    281: ("FDIST", 3, 3),
    282: ("FINV", 3, 3),
    283: ("FISHER", 1, 1),
    284: ("FISHERINV", 1, 1),
    285: ("FLOOR", 2, 2),
    286: ("GAMMADIST", 4, 4),
    287: ("GAMMAINV", 3, 3),
    288: ("CEILING", 2, 2),
    289: ("HYPGEOMDIST", 4, 4),
    290: ("LOGNORMDIST", 3, 3),
    291: ("LOGINV", 3, 3),
    292: ("NEGBINOMDIST", 3, 3),
    293: ("NORMDIST", 4, 4),
    294: ("NORMSDIST", 1, 1),
    295: ("NORMINV", 3, 3),
    296: ("NORMSINV", 1, 1),
    297: ("STANDARDIZE", 3, 3),
    298: ("ODD", 1, 1),
    299: ("PERMUT", 2, 2),
    300: ("POISSON", 3, 3),
    301: ("TDIST", 3, 3),
    302: ("WEIBULL", 4, 4),
    303: ("SUMXMY2", 2, 2),
    304: ("SUMX2MY2", 2, 2),
    305: ("SUMX2PY2", 2, 2),
    306: ("CHITEST", 2, 2),
    307: ("CORREL", 2, 2),
    308: ("COVAR", 2, 2),
    309: ("FORECAST", 3, 3),
    310: ("FTEST", 2, 2),
    311: ("INTERCEPT", 2, 2),
    312: ("PEARSON", 2, 2),
    313: ("RSQ", 2, 2),
    314: ("STEYX", 2, 2),
    315: ("SLOPE", 2, 2),
    316: ("TTEST", 4, 4),
    317: ("PROB", 3, 4),
    318: ("DEVSQ", 1, 255),
    319: ("GEOMEAN", 1, 255),
    320: ("HARMEAN", 1, 255),
    321: ("SUMSQ", 1, 255),
    322: ("KURT", 1, 255),
    323: ("SKEW", 1, 255),
    324: ("ZTEST", 2, 3),
    325: ("LARGE", 2, 2),
    326: ("SMALL", 2, 2),
    327: ("QUARTILE", 2, 2),
    328: ("PERCENTILE", 2, 2),
    329: ("PERCENTRANK", 2, 3),
    330: ("MODE", 1, 255),
    331: ("TRIMMEAN", 2, 2),
    332: ("TINV", 2, 2),
    334: ("MOVIE.COMMAND", 3, 4),
    335: ("GET.MOVIE", 2, 3),
    336: ("CONCATENATE", 1, 255),
    337: ("POWER", 2, 2),
    338: ("PIVOT.ADD.DATA", 2, 9),
    339: ("GET.PIVOT.TABLE", 1, 2),
    340: ("GET.PIVOT.FIELD", 1, 3),
    341: ("GET.PIVOT.ITEM", 1, 4),
    342: ("RADIANS", 1, 1),
    343: ("DEGREES", 1, 1),
    344: ("SUBTOTAL", 2, 255),
    345: ("SUMIF", 2, 3),
    346: ("COUNTIF", 2, 2),
    347: ("COUNTBLANK", 1, 1),
    348: ("SCENARIO.GET", 1, 2),
    349: ("OPTIONS.LISTS.GET", 1, 1),
    350: ("ISPMT", 4, 4),
    351: ("DATEDIF", 3, 3),
    352: ("DATESTRING", 1, 1),
    353: ("NUMBERSTRING", 2, 2),
    354: ("ROMAN", 1, 2),
    355: ("OPEN.DIALOG", 0, 4),
    356: ("SAVE.DIALOG", 0, 5),
    357: ("VIEW.GET", 1, 2),
    358: ("GETPIVOTDATA", 2, 254),
    359: ("HYPERLINK", 1, 2),
    360: ("PHONETIC", 1, 1),
    361: ("AVERAGEA", 1, 255),
    362: ("MAXA", 1, 255),
    363: ("MINA", 1, 255),
    364: ("STDEVPA", 1, 255),
    365: ("VARPA", 1, 255),
    366: ("STDEVA", 1, 255),
    367: ("VARA", 1, 255),
    368: ("BAHTTEXT", 1, 1),
    369: ("THAIDAYOFWEEK", 1, 1),
    370: ("THAIDIGIT", 1, 1),
    371: ("THAIMONTHOFYEAR", 1, 1),
    372: ("THAINUMSOUND", 1, 1),
    373: ("THAINUMSTRING", 1, 1),
    374: ("THAISTRINGLENGTH", 1, 1),
    375: ("ISTHAIDIGIT", 1, 1),
    376: ("ROUNDBAHTDOWN", 1, 1),
    377: ("ROUNDBAHTUP", 1, 1),
    378: ("THAIYEAR", 1, 1),
    379: ("RTD", 3, 255),
    380: ("CUBEVALUE", 1, 255),
    381: ("CUBEMEMBER", 2, 3),
    382: ("CUBEMEMBERPROPERTY", 3, 3),
    383: ("CUBERANKEDMEMBER", 3, 4),
    384: ("HEX2BIN", 1, 2),
    385: ("HEX2DEC", 1, 1),
    386: ("HEX2OCT", 1, 2),
    387: ("DEC2BIN", 1, 2),
    388: ("DEC2HEX", 1, 2),
    389: ("DEC2OCT", 1, 2),
    390: ("OCT2BIN", 1, 2),
    391: ("OCT2HEX", 1, 2),
    392: ("OCT2DEC", 1, 1),
    393: ("BIN2DEC", 1, 1),
    394: ("BIN2OCT", 1, 2),
    395: ("BIN2HEX", 1, 2),
    396: ("IMSUB", 2, 2),
    397: ("IMDIV", 2, 2),
    398: ("IMPOWER", 2, 2),
    399: ("IMABS", 1, 1),
    400: ("IMSQRT", 1, 1),
    401: ("IMLN", 1, 1),
    402: ("IMLOG2", 1, 1),
    403: ("IMLOG10", 1, 1),
    404: ("IMSIN", 1, 1),
    405: ("IMCOS", 1, 1),
    406: ("IMEXP", 1, 1),
    407: ("IMARGUMENT", 1, 1),
    408: ("IMCONJUGATE", 1, 1),
    409: ("IMAGINARY", 1, 1),
    410: ("IMREAL", 1, 1),
    411: ("COMPLEX", 2, 3),
    412: ("IMSUM", 1, 255),
    413: ("IMPRODUCT", 1, 255),
    414: ("SERIESSUM", 4, 4),
    415: ("FACTDOUBLE", 1, 1),
    416: ("SQRTPI", 1, 1),
    417: ("QUOTIENT", 2, 2),
    418: ("DELTA", 1, 2),
    419: ("GESTEP", 1, 2),
    420: ("ISEVEN", 1, 1),
    421: ("ISODD", 1, 1),
    422: ("MROUND", 2, 2),
    423: ("ERF", 1, 2),
    424: ("ERFC", 1, 1),
    425: ("BESSELJ", 2, 2),
    426: ("BESSELK", 2, 2),
    427: ("BESSELY", 2, 2),
    428: ("BESSELI", 2, 2),
    429: ("XIRR", 2, 3),
    430: ("XNPV", 3, 3),
    431: ("PRICEMAT", 5, 6),
    432: ("YIELDMAT", 5, 6),
    433: ("INTRATE", 4, 5),
    434: ("RECEIVED", 4, 5),
    435: ("DISC", 4, 5),
    436: ("PRICEDISC", 4, 5),
    437: ("YIELDDISC", 4, 5),
    438: ("TBILLEQ", 3, 3),
    439: ("TBILLPRICE", 3, 3),
    440: ("TBILLYIELD", 3, 3),
    441: ("PRICE", 6, 7),
    442: ("YIELD", 6, 7),
    443: ("DOLLARDE", 2, 2),
    444: ("DOLLARFR", 2, 2),
    445: ("NOMINAL", 2, 2),
    446: ("EFFECT", 2, 2),
    447: ("CUMPRINC", 6, 6),
    448: ("CUMIPMT", 6, 6),
    449: ("EDATE", 2, 2),
    450: ("EOMONTH", 2, 2),
    451: ("YEARFRAC", 2, 3),
    452: ("COUPDAYBS", 3, 4),
    453: ("COUPDAYS", 3, 4),
    454: ("COUPDAYSNC", 3, 4),
    455: ("COUPNCD", 3, 4),
    456: ("COUPNUM", 3, 4),
    457: ("COUPPCD", 3, 4),
    458: ("DURATION", 5, 6),
    459: ("MDURATION", 5, 6),
    460: ("ODDLPRICE", 7, 8),
    461: ("ODDLYIELD", 7, 8),
    462: ("ODDFPRICE", 8, 9),
    463: ("ODDFYIELD", 8, 9),
    464: ("RANDBETWEEN", 2, 2),
    465: ("WEEKNUM", 1, 2),
    466: ("AMORDEGRC", 6, 7),
    467: ("AMORLINC", 6, 7),
    469: ("ACCRINT", 6, 8),
    470: ("ACCRINTM", 4, 5),
    471: ("WORKDAY", 2, 3),
    472: ("NETWORKDAYS", 2, 3),
    473: ("GCD", 1, 255),
    474: ("MULTINOMIAL", 1, 255),
    475: ("LCM", 1, 255),
    476: ("FVSCHEDULE", 2, 2),
    477: ("CUBEKPIMEMBER", 3, 4),
    478: ("CUBESET", 2, 5),
    479: ("CUBESETCOUNT", 1, 1),
    480: ("IFERROR", 2, 2),
    481: ("COUNTIFS", 2, 254),
    482: ("SUMIFS", 3, 255),
    483: ("AVERAGEIF", 2, 3),
    484: ("AVERAGEIFS", 3, 255),
}

# BIFF8 permits these legacy aliases; later files use the normal Ftab entries.
FTAB[92] = ("SERIESSUM", 4, 4)
# Correct two copied parameter-rule labels in the published Ftab table.
FTAB[117] = ("EXACT", 2, 2)
FTAB[211] = ("LENB", 1, 1)
# Empty aggregates are accepted by Excel and can be emitted in older BIFF files.
for _function in (0, 4, 169, 183, 321, 336):
    _name, _minimum, _maximum = FTAB[_function]
    FTAB[_function] = (_name, 0, _maximum)
_FUNCTION_NAMES = {entry[0] for entry in FTAB.values()}
ERRORS = {
    0x00: "#NULL!", 0x07: "#DIV/0!", 0x0F: "#VALUE!", 0x17: "#REF!",
    0x1D: "#NAME?", 0x24: "#NUM!", 0x2A: "#N/A", 0x2B: "#GETTING_DATA",
}
_BINARY = {
    0x03: ("+", 30), 0x04: ("-", 30), 0x05: ("*", 40), 0x06: ("/", 40),
    0x07: ("^", 50), 0x08: ("&", 20), 0x09: ("<", 10), 0x0A: ("<=", 10),
    0x0B: ("=", 10), 0x0C: (">=", 10), 0x0D: (">", 10), 0x0E: ("<>", 10),
    0x0F: (" ", 90), 0x10: (",", 80), 0x11: (":", 100),
}


class _Reader:
    def __init__(self, data):
        self.data = bytes(data)
        self.pos = 0

    def take(self, size):
        if size < 0 or self.pos + size > len(self.data):
            raise FormulaError("数式のトークンが途中で切れています")
        value = self.data[self.pos:self.pos + size]
        self.pos += size
        return value

    def unpack(self, code):
        return struct.unpack("<" + code, self.take(struct.calcsize("<" + code)))

    def uint8(self):
        return self.unpack("B")[0]

    def uint16(self):
        return self.unpack("H")[0]

    def uint32(self):
        return self.unpack("I")[0]

    def text(self, count, wide=True):
        try:
            return self.take(count * (2 if wide else 1)).decode("utf-16le" if wide else "latin1")
        except UnicodeDecodeError:
            raise FormulaError("数式の文字列が不正です") from None


@dataclass(frozen=True)
class _Expression:
    text: str
    precedence: int = 110
    operator: str = ""
    name: bool = False


def _number(number):
    if not math.isfinite(number):
        raise FormulaError("数式に有限でない数値があります")
    if number == 0:
        return "0"
    # Excel rounds formula constants before choosing fixed/exponential spelling.
    rounded = Decimal(format(number, ".15g"))
    fixed = format(rounded, "f")
    if "." in fixed:
        fixed = fixed.rstrip("0").rstrip(".")
    if len(fixed) <= 21:
        return fixed
    mantissa, exponent = format(rounded, "E").split("E")
    if "." in mantissa:
        mantissa = mantissa.rstrip("0").rstrip(".")
    return mantissa + "E" + format(int(exponent), "+03d")


def _quote(value):
    return '"' + value.replace('"', '""') + '"'


def _error(code):
    if code not in ERRORS:
        raise FormulaError("数式に未知のエラー値があります")
    return ERRORS[code]


def _boolean(value):
    if value not in (0, 1):
        raise FormulaError("数式の真偽値が不正です")
    return "TRUE" if value else "FALSE"


def _column(column):
    result = ""
    column += 1
    while column:
        column, rest = divmod(column - 1, 26)
        result = chr(65 + rest) + result
    return result


def quote_sheet_name(value):
    """Use Excel's unquoted spelling when it cannot resemble a reference."""
    if not isinstance(value, str) or not value:
        raise FormulaError("数式のシート参照を解決できません")
    simple = not value[0].isdigit() and all(
        char.isalpha() or char.isdigit() or char in "_." for char in value)
    a1 = re.fullmatch(r"([A-Za-z]{1,3})([1-9][0-9]{0,6})", value)
    reference = False
    if a1:
        column = 0
        for char in a1[1].upper():
            column = column * 26 + ord(char) - ord("A") + 1
        reference = column <= 16384 and int(a1[2]) <= 1048576
    r1c1 = re.fullmatch(r"R[0-9]*(?:C[0-9]*)?|C[0-9]*", value, re.IGNORECASE)
    if simple and not reference and not r1c1 and value.upper() not in ("TRUE", "FALSE"):
        return value
    return "'" + value.replace("'", "''") + "'"


def _lookup(values, index, description):
    try:
        value = values[index]
    except (KeyError, IndexError, TypeError):
        raise FormulaError("数式の" + description + "を解決できません") from None
    if value is None:
        raise FormulaError("外部ブックを参照する数式は扱えません")
    return value


def _sheet(values, index):
    value = _lookup(values, index, "シート参照")
    if isinstance(value, (tuple, list)) and len(value) == 2:
        first, last = value
        if first == last:
            value = first
        else:
            # The single pair of quotes encloses both ends of a 3-D range.
            first_text, last_text = quote_sheet_name(first), quote_sheet_name(last)
            if first_text.startswith("'") or last_text.startswith("'"):
                return "'" + (first + ":" + last).replace("'", "''") + "'!"
            return first + ":" + last + "!"
    return quote_sheet_name(value) + "!"


def _name(value):
    if not isinstance(value, str) or not value:
        raise FormulaError("数式の名前定義が不正です")
    if value.lower().startswith("_xlfn.") and value[6:].upper() in _FUNCTION_NAMES:
        return value[6:]
    return value


def _ref_parts(reader, version, row, col, relative=False):
    target_row = reader.uint32() if version == 12 else reader.uint16()
    flags = reader.uint16()
    row_relative, col_relative = bool(flags & 0x8000), bool(flags & 0x4000)
    max_rows, max_columns = (1048576, 16384) if version == 12 else (65536, 256)
    if version == 8 and (not relative or not col_relative) and flags & 0x3FFF >= 256:
        raise FormulaError("数式の参照先が Excel の範囲外です")
    target_col = flags & (max_columns - 1)
    if relative:
        # The modular arithmetic also handles the signed negative encoding.
        if row_relative:
            target_row = (row + target_row) % max_rows
        if col_relative:
            target_col = (col + target_col) % max_columns
    if target_row >= max_rows or target_col >= max_columns:
        raise FormulaError("数式の参照先が Excel の範囲外です")
    return target_row, target_col, row_relative, col_relative


def _address(parts, rows=True, columns=True):
    row, col, row_relative, col_relative = parts
    return (("" if col_relative else "$") + _column(col) if columns else "") + (
        ("" if row_relative else "$") + str(row + 1) if rows else "")


def _reference(reader, version, row, col, relative=False, area=False):
    if not area:
        return _address(_ref_parts(reader, version, row, col, relative))
    # Area structures group both rows before the two column/flag fields.
    row_code = "I" if version == 12 else "H"
    first_row, last_row = reader.unpack(row_code * 2)
    first_col, last_col = reader.unpack("HH")
    first = _Reader(struct.pack("<" + row_code + "H", first_row, first_col))
    last = _Reader(struct.pack("<" + row_code + "H", last_row, last_col))
    first = _ref_parts(first, version, row, col, relative)
    last = _ref_parts(last, version, row, col, relative)
    max_row, max_col = (1048575, 16383) if version == 12 else (65535, 255)
    columns = not (first[1] == 0 and last[1] == max_col)
    rows = not (first[0] == 0 and last[0] == max_row)
    if not columns:
        rows = True
    return _address(first, rows, columns) + ":" + _address(last, rows, columns)


def _array(reader, version):
    if version == 12:
        rows, columns = reader.unpack("II")
    else:
        columns, rows = reader.unpack("BH")
        rows, columns = rows + 1, columns + 1
    # Each value needs at least a type byte and a payload byte. Reject an
    # impossible product before allocating or looping over attacker input.
    if not rows or not columns or rows * columns > (len(reader.data) - reader.pos) // 2:
        raise FormulaError("数式の配列定数の大きさが不正です")
    result = []
    for _row in range(rows):
        values = []
        for _col in range(columns):
            kind = reader.uint8()
            if kind == (0 if version == 12 else 1):
                text = _number(reader.unpack("d")[0])
            elif kind == (1 if version == 12 else 2):
                count = reader.uint16()
                wide = True if version == 12 else bool(reader.uint8() & 1)
                text = _quote(reader.text(count, wide))
            elif kind == (2 if version == 12 else 4):
                text = _boolean(reader.uint8())
                if version != 12:
                    reader.take(7)
            elif kind == (4 if version == 12 else 16):
                text = _error(reader.uint8())
                reader.take(3 if version == 12 else 7)
            elif version != 12 and kind == 0:
                reader.take(8)
                text = ""
            else:
                raise FormulaError("数式の配列定数の型を扱えません")
            values.append(text)
        result.append(",".join(values))
    return "{" + ";".join(result) + "}"


def decode_formula(rgce, rgcb=b"", *, version=12, row=0, col=0,
                   sheets=None, names=None, external_names=None):
    """Decode rgce/rgcb; row and col are the zero-based formula cell position.

    sheets maps ixti to a local sheet name or a pair of endpoint names. names
    maps one-based IDs to names. external_names maps (ixti, one-based ID) to
    local add-in names; callers must exclude all external-workbook names.
    PtgExp is deliberately left to the workbook reader's shared/array lookup.
    """
    if version in (8, 80):
        version = 8
    elif version != 12:
        raise FormulaError("対応していない数式の BIFF 版です")
    reader, extra = _Reader(rgce), _Reader(rgcb)
    stack = []
    space = ""
    paren_before, paren_after = "", ""

    def pop(count):
        if count < 0 or len(stack) < count:
            raise FormulaError("数式の演算子と引数の数が合いません")
        if not count:
            return []
        result = stack[-count:]
        del stack[-count:]
        return result

    def append(text, precedence=110, operator="", name=False):
        nonlocal space
        stack.append(_Expression(space + text, precedence, operator, name))
        space = ""

    def argument(expression):
        return "(" + expression.text + ")" if expression.operator == "," else expression.text

    while reader.pos < len(reader.data):
        token = reader.uint8()
        if token & 0x80:
            raise FormulaError("未知の数式トークンです: 0x%02x" % token)
        ptg = (token & 0x1F) | 0x20 if token >= 0x20 else token
        if ptg in _BINARY:
            left, right = pop(2)
            operator, precedence = _BINARY[ptg]
            left_text = "(" + left.text + ")" if left.precedence < precedence else left.text
            right_text = "(" + right.text + ")" if right.precedence <= precedence else right.text
            # PtgAttrSpace before a binary token belongs before its operator.
            stack.append(_Expression(left_text + space + operator + right_text, precedence, operator))
            space = ""
        elif ptg in (0x12, 0x13, 0x14):
            operand, = pop(1)
            precedence = 60 if ptg == 0x14 else 70
            text = "(" + operand.text + ")" if operand.precedence < precedence else operand.text
            append(text + "%" if ptg == 0x14 else ("+" if ptg == 0x12 else "-") + text,
                   precedence, "%" if ptg == 0x14 else "unary")
        elif ptg == 0x15:
            operand, = pop(1)
            append(paren_before + "(" + operand.text + paren_after + ")")
            paren_before, paren_after = "", ""
        elif ptg == 0x16:
            append("")
        elif ptg == 0x17:
            if version == 12:
                text = reader.text(reader.uint16())
            else:
                count, flags = reader.unpack("BB")
                if flags & ~1:
                    raise FormulaError("数式の文字列のフラグが不正です")
                text = reader.text(count, bool(flags & 1))
            append(_quote(text))
        elif ptg == 0x19:
            flags, value = reader.unpack("BH")
            if flags & 0x40:
                kind, count = value & 0xFF, value >> 8
                if kind > 6:
                    raise FormulaError("数式の空白トークンが不正です")
                text = ("\r" if kind in (1, 3, 5) else " ") * count
                if kind in (2, 3):
                    paren_before += text
                elif kind in (4, 5):
                    paren_after += text
                else:
                    space += text
            elif flags & 0x10:
                operand, = pop(1)
                append("SUM(" + argument(operand) + ")")
            elif flags & 0x04:
                reader.take((value + 1) * 2)
            # If, GoTo, Semi, Baxcel and IfError are execution hints.
            elif flags not in (0, 1, 2, 8, 0x20, 0x21, 0x80):
                raise FormulaError("未知の数式属性トークンです")
        elif ptg == 0x1C:
            append(_error(reader.uint8()))
        elif ptg == 0x1D:
            append(_boolean(reader.uint8()))
        elif ptg == 0x1E:
            append(str(reader.uint16()))
        elif ptg == 0x1F:
            append(_number(reader.unpack("d")[0]))
        elif ptg == 0x20:
            reader.take(14 if version == 12 else 7)
            append(_array(extra, version))
        elif ptg in (0x21, 0x22):
            count = reader.uint8() if ptg == 0x22 else None
            function = reader.uint16()
            if function & 0x8000:
                raise FormulaError("マクロ関数の数式は扱えません")
            if function == 255 and count is not None:
                args = pop(count)
                if not args or not args[0].name:
                    raise FormulaError("アドイン関数の名前を解決できません")
                name, args = args[0].text, args[1:]
            else:
                if function not in FTAB:
                    raise FormulaError("未知の数式関数です: %d" % function)
                name, minimum, maximum = FTAB[function]
                if count is None:
                    if minimum != maximum:
                        raise FormulaError("可変長の関数に引数の数がありません")
                    count = minimum
                if not minimum <= count <= maximum:
                    raise FormulaError("数式の関数の引数の数が不正です")
                args = pop(count)
            append(name + "(" + ",".join(argument(arg) for arg in args) + ")")
        elif ptg == 0x23:
            append(_name(_lookup(names, reader.uint32(), "名前定義")), name=True)
        elif ptg in (0x24, 0x25, 0x2C, 0x2D):
            append(_reference(reader, version, row, col, ptg in (0x2C, 0x2D), ptg in (0x25, 0x2D)))
        elif ptg in (0x26, 0x27, 0x28):
            reader.take(4)
            reader.take(2)
            if ptg == 0x26:
                count = extra.uint32() if version == 12 else extra.uint16()
                extra.take(count * (16 if version == 12 else 8))
        elif ptg == 0x29:
            reader.take(2)
        elif ptg in (0x2A, 0x2B):
            reader.take((6 if version == 12 else 4) * (2 if ptg == 0x2B else 1))
            append("#REF!")
        elif ptg == 0x39:
            ixti, index = reader.unpack("HI")
            try:
                value = external_names[(ixti, index)]
            except (KeyError, TypeError):
                value = _lookup(external_names, index, "アドイン関数の名前")
            if value is None:
                raise FormulaError("外部ブックを参照する数式は扱えません")
            append(_name(value), name=True)
        elif ptg in (0x3A, 0x3B):
            prefix = _sheet(sheets, reader.uint16())
            append(prefix + _reference(reader, version, row, col, area=ptg == 0x3B))
        elif ptg in (0x3C, 0x3D):
            prefix = _sheet(sheets, reader.uint16())
            reader.take((6 if version == 12 else 4) * (2 if ptg == 0x3D else 1))
            append(prefix + "#REF!")
        elif ptg == 0x01:
            raise FormulaError("共有数式・配列数式の参照は呼び出し側で解決してください")
        else:
            raise FormulaError("対応していない数式トークンです: 0x%02x" % token)
    if len(stack) != 1 or paren_before or paren_after:
        raise FormulaError("数式のトークンの並びが不正です")
    if extra.pos != len(extra.data):
        raise FormulaError("数式の追加データが余っています")
    return space + stack[0].text
