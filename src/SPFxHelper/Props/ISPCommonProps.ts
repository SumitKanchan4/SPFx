interface IDocResponse {
    ok: boolean;
    imageUrl?: string;
    fileName?: string;
    error: Error;
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

export { IDocResponse };
export { IUserProps };