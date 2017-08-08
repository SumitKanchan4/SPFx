interface ILogger {
    errorMethod: string;
    errorMessage: string;
}

interface ILoggerResponse {
    ok: boolean;
    status: number;
    statusText: string;
    success: boolean;
    errorMethod: string;
}

enum errorType {
    ERROR,
    DEBUG
}

export { errorType };
export { ILogger };
export { ILoggerResponse };