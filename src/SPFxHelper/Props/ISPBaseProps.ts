interface ISPPostRequest {
    url: string;
    body: string;
}

interface ISPBaseResponse {
    result?: any;
    ok: boolean;
    error?: Error;
}

export { ISPBaseResponse };
export { ISPPostRequest };