export default interface UserData {
    employee: string | ISiteUserInfo;
    reward: string;
}

interface ISiteUserInfo {
    Email: string;
    Title: string;
}