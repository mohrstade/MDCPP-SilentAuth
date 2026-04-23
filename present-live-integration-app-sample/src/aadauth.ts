import {
    AccountInfo,
    PopupRequest,
    PublicClientApplication,
    SilentRequest,
} from "@azure/msal-browser";
import { trySilentAuth } from "@microsoft/document-collaboration-sdk";

const aadClientId = process.env.ENTRA_APPID ?? "";
if (aadClientId === "") {
    throw new Error("Test app Entra client ID is missing. Please ensure it's defined in the .env file.");
}

const redirectUri = window.location.pathname;
const defaultScopes = ["User.Read", "Files.Read"];
const LAST_ACCOUNT_ID_KEY = "team-startpage:lastAccountId";
const LAST_USERNAME_KEY = "team-startpage:lastUsername";
const SILENT_AUTH_PLATFORM = "Web";

const app = new PublicClientApplication({
    auth: {
        authority: "https://login.microsoftonline.com/common",
        clientId: aadClientId,
        redirectUri,
    },
    cache: {
        cacheLocation: "localStorage",
    },
});

let activeAccount: AccountInfo | null = null;
let initialized = false;

export function combine(...paths: string[]) {
    return paths
        .map(path => path.replace(/^[\\|/]/, "").replace(/[\\|/]$/, ""))
        .join("/")
        .replace(/\\/g, "/");
}

function setActiveAccount(account: AccountInfo | null) {
    activeAccount = account;
    if (account) {
        app.setActiveAccount(account);
        window.localStorage.setItem(LAST_ACCOUNT_ID_KEY, account.homeAccountId);
        if (account.username) {
            window.localStorage.setItem(LAST_USERNAME_KEY, account.username);
        }
    } else {
        app.setActiveAccount(null);
        window.localStorage.removeItem(LAST_ACCOUNT_ID_KEY);
        window.localStorage.removeItem(LAST_USERNAME_KEY);
    }
}

function clearRuntimeAccount() {
    activeAccount = null;
    app.setActiveAccount(null);
}

function getPreferredCachedAccount() {
    const accounts = app.getAllAccounts();
    if (!accounts.length) return null;

    const lastAccountId = window.localStorage.getItem(LAST_ACCOUNT_ID_KEY);
    const lastUsername = window.localStorage.getItem(LAST_USERNAME_KEY);

    return (
        accounts.find(account => account.homeAccountId === lastAccountId) ||
        accounts.find(account => account.username === lastUsername) ||
        app.getActiveAccount() ||
        accounts[0] ||
        null
    );
}

async function attemptMdcppSilentAuth(loginHint?: string | null) {
    if (!loginHint) return false;

    try {
        const success = await trySilentAuth(aadClientId, loginHint, SILENT_AUTH_PLATFORM);
        if (!success) {
            return false;
        }

        const cachedAccount = getPreferredCachedAccount();
        if (cachedAccount) {
            setActiveAccount(cachedAccount);
            return true;
        }

        return false;
    } catch {
        return false;
    }
}

export async function initializeAuth() {
    if (!initialized) {
        await app.initialize();
        const response = await app.handleRedirectPromise();
        if (response?.account) {
            setActiveAccount(response.account);
        } else {
            const cachedAccount = getPreferredCachedAccount();
            if (cachedAccount) {
                setActiveAccount(cachedAccount);
            } else {
                const loginHint = window.localStorage.getItem(LAST_USERNAME_KEY);
                const silentAuthSucceeded = await attemptMdcppSilentAuth(loginHint);
                if (!silentAuthSucceeded) {
                    clearRuntimeAccount();
                }
            }
        }
        initialized = true;
    }

    return activeAccount;
}

export function getActiveAccount() {
    return activeAccount;
}

export function isSignedIn() {
    return activeAccount !== null;
}

export async function signIn() {
    const loginHint = window.localStorage.getItem(LAST_USERNAME_KEY);
    const silentAuthSucceeded = await attemptMdcppSilentAuth(loginHint);
    if (silentAuthSucceeded && activeAccount) {
        return activeAccount;
    }

    const request: PopupRequest = {
        prompt: "login",
        scopes: defaultScopes,
    };

    const response = await app.loginPopup(request);
    setActiveAccount(response.account);
    return response.account;
}

export async function signOut() {
    if (!activeAccount) return;

    await app.logoutPopup({
        account: activeAccount,
        mainWindowRedirectUri: window.location.origin + window.location.pathname,
    });

    setActiveAccount(null);
}

function normalizeScopes(scopes: string[]) {
    const merged = new Set<string>([...defaultScopes, ...scopes]);
    return Array.from(merged);
}

export async function getGraphToken(scopes: string[] = defaultScopes) {
    if (!activeAccount) {
        throw new Error("User is not signed in.");
    }

    const resolvedScopes = normalizeScopes(scopes);
    const silentRequest: SilentRequest = {
        account: activeAccount,
        scopes: resolvedScopes,
    };

    try {
        const response = await app.acquireTokenSilent(silentRequest);
        return response.accessToken;
    } catch (error: any) {
        const message = error?.message || "";
        if (message.toLowerCase().includes("interaction")) {
            const loginHint = activeAccount?.username || window.localStorage.getItem(LAST_USERNAME_KEY);
            const silentAuthSucceeded = await attemptMdcppSilentAuth(loginHint);
            if (silentAuthSucceeded && activeAccount) {
                const retryResponse = await app.acquireTokenSilent({
                    account: activeAccount,
                    scopes: resolvedScopes,
                });
                return retryResponse.accessToken;
            }
            throw new Error("Your Microsoft session needs attention or new permissions must be accepted. Please reconnect when you are ready.");
        }
        throw error;
    }
}

export async function getToken(command: { resource: string; type: string }) {
    if (command.type === "Default") {
        return getGraphToken(defaultScopes);
    }

    if (command.type === "SharePoint" || command.type === "SharePoint_SelfIssued") {
        return getGraphToken([`${combine(command.resource, ".default")}`]);
    }

    return "";
}

export async function graphGet<T>(url: string, scopes: string[] = defaultScopes): Promise<T> {
    return graphRequest<T>(url, {
        method: "GET",
    }, scopes);
}

export async function graphPost<T>(url: string, body: Record<string, unknown> = {}, scopes: string[] = defaultScopes): Promise<T> {
    return graphRequest<T>(url, {
        method: "POST",
        headers: {
            "Content-Type": "application/json",
        },
        body: JSON.stringify(body),
    }, scopes);
}

async function graphRequest<T>(url: string, init: RequestInit, scopes: string[] = defaultScopes): Promise<T> {
    const token = await getGraphToken(scopes);
    const response = await fetch(url, {
        ...init,
        headers: {
            Authorization: `Bearer ${token}`,
            ...(init.headers || {}),
        },
    });

    const data = await response.json().catch(() => ({}));
    if (!response.ok) {
        const errorMessage = (data as any)?.error?.message || response.statusText || "Graph request failed";
        throw new Error(errorMessage);
    }

    return data as T;
}
