interface ISPBaseResponse {
    ok: boolean;
    status: number;
    statusText: string;
    result: any;
    errorMethod: string;
    responseJSON?: string;
}

interface ISPPostRequest {
    url: string;
    body: string;
}

export { ISPBaseResponse };
export { ISPPostRequest };