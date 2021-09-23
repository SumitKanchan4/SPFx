interface ISPPostRequest {
    url: string;
    body: string;
}

interface ISPBaseResponse {
    result?: any;
    ok: boolean;
    error?: Error;
    status?: number;
    statusText?: string;
    errorMethod?: string;
    responseJSON?: string;
}

export { ISPBaseResponse };
export { ISPPostRequest };