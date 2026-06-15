export default async function request(url, options = {}) {
    const {
        method = 'GET',
        timeoutMs = 120000,
        ...fetchOptions
    } = options;
    const upperMethod = method.toUpperCase();
    const canHaveBody = !['GET', 'HEAD'].includes(upperMethod);

    let res;

    try {
        res = await fetch(url, {
            ...fetchOptions,
            method,
            signal: AbortSignal.timeout(timeoutMs),
            headers: {
                'Content-Type': 'application/json',
                ...fetchOptions.headers,
            },
            ...(canHaveBody ? {
                body: JSON.stringify(fetchOptions.body || {})
            } : {}),
        });
    } catch (error) {
        if (error.name === "TimeoutError") {
            throw new Error(`请求超时，豆包在 ${Math.round(timeoutMs / 1000)} 秒内没有返回结果`);
        }

        throw error;
    }

    const data = await res.json();
    if (!res.ok) {
        const message = data?.error?.message
            || data?.message
            || `请求失败：${res.status}`;

        throw new Error(message);
    }

    return data;
}
