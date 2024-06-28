export class Warning extends Error { 
    public isWarning: boolean = true;
    constructor(message: string) {
      super(message);
    }
}