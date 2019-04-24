import { ServiceScope, Log } from '@microsoft/sp-core-library';

/**
 * Generic Logger
 */
class SPLogger {

    /**
    * Returns the caller method name, assuming it is third in error stack
    */
    private static get getMethodName(): string {
        var stackTrace = (new Error()).stack; // Only tested in latest FF and Chrome
        let stack: any = stackTrace.replace(/^Error\s+/, ''); // Sanitize Chrome
        let callerName: string = stack.split("\n")[2]; // 1st item is this, 3nd item is original caller
        callerName = callerName.replace(/^\s+at Object./, ''); // Sanitize Chrome
        callerName = callerName.replace(/ \(.+\)$/, ''); // Sanitize Chrome
        callerName = callerName.replace(/\@.+/, '');
        callerName = callerName.replace("prototype.", "");
        callerName = callerName.split("at").length > 1 ? callerName.split("at")[1] : callerName;
        return callerName.trim();
    }

    /**
     * Logs error to the console
     * @param message : Error object
     * @param scope : the service scope that the source uses. A service scope can provide more context information (e.g., web part information) to the logged error.
     */
    public static logError(message: Error, scope?: ServiceScope): void {

        Log.error(this.getMethodName, message, scope);
    }

    /**
     * Logs a general informational message.
     * @param message the message to be logged
     * @param scope the service scope that the source uses. A service scope can provide more context information (e.g., web part information) to the logged error.
     */
    public static logInfo(message: string, scope?: ServiceScope): void {
        Log.info(this.getMethodName, message, scope);
    }

    /**
     * Logs a message which contains detailed information that is generally only needed for troubleshooting.
     * @param message the message to be logged
     * @param scope the service scope that the source uses. A service scope can provide more context information (e.g., web part information) to the logged error.
     */
    public static logVerbose(message: string, scope?: ServiceScope): void {
        Log.verbose(this.getMethodName, message, scope);
    }

    /**
     * Logs a warning.
     * @param message the message to be logged
     * @param scope the service scope that the source uses. A service scope can provide more context information (e.g., web part information) to the logged error.
     */
    public static logWarning(message: string, scope?: ServiceScope): void {
        Log.warn(this.getMethodName, message, scope);
    }
}

export { SPLogger };