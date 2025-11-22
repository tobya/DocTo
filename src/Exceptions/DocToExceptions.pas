unit DocToExceptions;

interface
uses Classes,   ActiveX, ComObj, WinINet, Variants, sysutils, Types, StrUtils, TypInfo;

type

          EDocToException = class(Exception);
          ESheetIndexOutOfBounds = class(EDocToException);
          EInvalidParameterCombination = class(EDocToException);

implementation

end.
