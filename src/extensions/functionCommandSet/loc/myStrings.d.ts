declare interface IFunctionCommandSetStrings {
  Command1: string;
  Command2: string;
    Command3: string;
}

declare module 'functionCommandSetStrings' {
  const strings: IFunctionCommandSetStrings;
  export = strings;
}
