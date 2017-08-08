interface IDocResponse {
    ok: boolean;
    status: number;
    statusText: string;
    image: string;
    fileName: string;
    fileUrl: string;
    id:string;
    success:boolean;
    errorMethod: string;
}

interface IDoc {
    fileName: string;
    fileUrl: string;
    id: string;
}

interface IUserProps {
    id: number;
    name: string;
    imageUrl: string;
    email: string;
    department: string;
    jobTitle: string;
    workPhone: string;
    userName: string;
}

export { IDoc };
export { IDocResponse };
export { IUserProps };