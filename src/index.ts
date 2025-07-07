/**
 * Cloudflare Worker for Spotify Liked Songs Shuffle with Last.fm Play Counts (TypeScript)
 *
 * This worker authenticates with Spotify, fetches a user's liked songs,
 * retrieves their play counts from Last.fm, and creates a new Spotify playlist
 * where songs are cumulatively shuffled based on their play count.
 * (e.g., first block shuffles songs played once, second block shuffles songs
 * played once AND twice, etc.).
 *
 * Environment Variables Required:
 * - SPOTIFY_CLIENT_ID: Your Spotify Application Client ID
 * - SPOTIFY_CLIENT_SECRET: Your Spotify Application Client Secret
 * - SPOTIFY_REDIRECT_URI: The redirect URI configured in your Spotify App (e.g., https://your-worker-domain.workers.dev/callback)
 * - LASTFM_API_KEY: Your Last.fm API Key
 * - LASTFM_USERNAME: Your Last.fm Username (for fetching your specific play counts)
 *
 * KV Namespace Required:
 * - SPOTIFY_TOKENS: A KV namespace to store user-specific Spotify access and refresh tokens.
 */

// Declare global types for Cloudflare Worker environment variables and KV namespace
// This is necessary for TypeScript to recognize these global variables.
declare const SPOTIFY_CLIENT_ID: string;
declare const SPOTIFY_CLIENT_SECRET: string;
declare const SPOTIFY_REDIRECT_URI: string;
declare const LASTFM_API_KEY: string;
declare const LASTFM_USERNAME: string;
declare const SPOTIFY_TOKENS: KVNamespace;

// Define constants for Spotify API endpoints
const SPOTIFY_AUTH_URL: string = 'https://accounts.spotify.com/authorize';
const SPOTIFY_TOKEN_URL: string = 'https://accounts.spotify.com/api/token';
const SPOTIFY_API_BASE_URL: string = 'https://api.spotify.com/v1';

// Define constants for Last.fm API endpoints
const LASTFM_API_BASE_URL: string = 'http://ws.audioscrobbler.com/2.0/';

// --- Type Definitions ---

interface SpotifyTrack {
    id: string;
    name: string;
    artists: Array<{ name: string }>;
    uri: string;
    // Add other properties you might use from Spotify API if needed
    // e.g., album: { release_date: string }, popularity: number
    [key: string]: any; // Allow for other properties not explicitly defined
}

interface TrackWithPlayCount extends SpotifyTrack {
    lastFmPlayCount: number | null;
}

interface SpotifyTokenData {
    access_token: string;
    refresh_token: string;
    expires_at: number; // Unix timestamp in milliseconds
}

// --- Helper Functions ---

/**
 * Generates a random string for the state parameter in OAuth.
 * @param {number} length - The length of the random string.
 * @returns {string} A random string.
 */
function generateRandomString(length: number): string {
    let text: string = '';
    const possible: string = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
    for (let i = 0; i < length; i++) {
        text += possible.charAt(Math.floor(Math.random() * possible.length));
    }
    return text;
}

/**
 * Encodes an object into a URL-encoded query string.
 * @param {Record<string, string | number | boolean>} obj - The object to encode.
 * @returns {string} The URL-encoded query string.
 */
function encodeQueryParams(obj: Record<string, string | number | boolean>): string {
    return Object.entries(obj)
        .map(([key, value]) => `${encodeURIComponent(key)}=${encodeURIComponent(value)}`)
        .join('&');
}

/**
 * Fetches data from a Spotify API endpoint.
 * Handles authentication and error responses.
 * @param {string} url - The URL to fetch.
 * @param {string} accessToken - The Spotify access token.
 * @param {RequestInit} [options={}] - Fetch options.
 * @returns {Promise<any>} The JSON response from the API.
 * @throws {Error} If the API request fails.
 */
async function spotifyApiFetch(url: string, accessToken: string, options: RequestInit = {}): Promise<any> {
    const response: Response = await fetch(url, {
        headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/json',
            ...(options.headers as Record<string, string>), // Cast to Record<string, string> for spread
        },
        ...options,
    });

    if (!response.ok) {
        const errorData: any = await response.json().catch(() => ({ message: 'Unknown error' }));
        console.error(`Spotify API Error (${response.status}):`, errorData);
        throw new Error(`Spotify API request failed: ${response.status} - ${errorData.error?.message || errorData.message}`);
    }

    return response.json();
}

/**
 * Refreshes the Spotify access token using the refresh token.
 * Stores the new tokens in KV.
 * @param {string} userId - The user ID associated with the tokens in KV.
 * @param {string} refreshToken - The refresh token.
 * @returns {Promise<string>} The new access token.
 * @throws {Error} If token refresh fails.
 */
async function refreshSpotifyToken(userId: string, refreshToken: string): Promise<string> {
    const client_id: string = SPOTIFY_CLIENT_ID;
    const client_secret: string = SPOTIFY_CLIENT_SECRET;

    const authHeader: string = btoa(`${client_id}:${client_secret}`);

    const response: Response = await fetch(SPOTIFY_TOKEN_URL, {
        method: 'POST',
        headers: {
            'Authorization': `Basic ${authHeader}`,
            'Content-Type': 'application/x-www-form-urlencoded',
        },
        body: encodeQueryParams({
            grant_type: 'refresh_token',
            refresh_token: refreshToken,
        }),
    });

    if (!response.ok) {
        const errorData: any = await response.json().catch(() => ({ message: 'Unknown error' }));
        console.error('Token refresh failed:', errorData);
        throw new Error(`Failed to refresh Spotify token: ${response.status} - ${errorData.error_description || errorData.error}`);
    }

    const data: any = await response.json();
    const newAccessToken: string = data.access_token;
    const newRefreshToken: string = data.refresh_token || refreshToken; // Refresh token might not always be new

    // Store updated tokens in KV
    await SPOTIFY_TOKENS.put(userId, JSON.stringify({
        access_token: newAccessToken,
        refresh_token: newRefreshToken,
        expires_at: Date.now() + (data.expires_in * 1000)
    } as SpotifyTokenData));

    return newAccessToken;
}

/**
 * Retrieves valid Spotify access token for a user.
 * Refreshes if expired.
 * @param {string} userId - The user ID.
 * @returns {Promise<string>} The valid access token.
 * @throws {Error} If no valid token can be obtained.
 */
async function getValidAccessToken(userId: string): Promise<string> {
    const tokenDataStr: string | null = await SPOTIFY_TOKENS.get(userId);
    if (!tokenDataStr) {
        throw new Error('No Spotify tokens found for this user. Please log in first.');
    }

    let tokenData: SpotifyTokenData = JSON.parse(tokenDataStr);
    let accessToken: string = tokenData.access_token;

    // Check if token is expired (give a 5-minute buffer)
    if (tokenData.expires_at < Date.now() + (5 * 60 * 1000)) {
        console.log('Access token expired or near expiration. Refreshing...');
        accessToken = await refreshSpotifyToken(userId, tokenData.refresh_token);
    }
    return accessToken;
}

/**
 * Fetches all liked songs (tracks) for the authenticated user.
 * Handles pagination.
 * @param {string} accessToken - The Spotify access token.
 * @returns {Promise<SpotifyTrack[]>} An array of track objects.
 */
async function fetchAllLikedSongs(accessToken: string): Promise<SpotifyTrack[]> {
    let allTracks: SpotifyTrack[] = [];
    let nextUrl: string | null = `${SPOTIFY_API_BASE_URL}/me/tracks?limit=50`; // Max limit is 50

    while (nextUrl) {
        console.log(`Fetching liked songs from: ${nextUrl}`);
        const data: any = await spotifyApiFetch(nextUrl, accessToken);
        allTracks = allTracks.concat(data.items.map((item: any) => item.track as SpotifyTrack));
        nextUrl = data.next;
    }
    return allTracks;
}

/**
 * Fetches the play count for a specific track from Last.fm.
 * @param {string} artistName - The artist's name.
 * @param {string} trackName - The track's name.
 * @param {string} lastFmApiKey - Your Last.fm API key.
 * @param {string} lastFmUsername - Your Last.fm username.
 * @returns {Promise<number|null>} The play count, or null if not found/error.
 */
async function fetchLastFmPlayCount(artistName: string, trackName: string, lastFmApiKey: string, lastFmUsername: string): Promise<number | null> {
    const params: Record<string, string> = {
        method: 'track.getInfo',
        api_key: lastFmApiKey,
        artist: artistName,
        track: trackName,
        username: lastFmUsername, // Crucial for getting user-specific play count
        format: 'json'
    };
    const url: string = `${LASTFM_API_BASE_URL}?${encodeQueryParams(params)}`;

    try {
        // Add a small delay to avoid hitting Last.fm rate limits too quickly
        await new Promise(resolve => setTimeout(resolve, 100)); // 100ms delay

        const response: Response = await fetch(url);
        if (!response.ok) {
            console.warn(`Last.fm API Error (${response.status}) for ${artistName} - ${trackName}`);
            return null;
        }
        const data: any = await response.json();
        if (data && data.track && typeof data.track.userplaycount !== 'undefined') {
            return parseInt(data.track.userplaycount, 10);
        }
        console.log(`No Last.fm play count found for: ${artistName} - ${trackName}`);
        return null;
    } catch (error: any) {
        console.error(`Error fetching Last.fm play count for ${artistName} - ${trackName}:`, error);
        return null;
    }
}

/**
 * Implements the custom shuffle logic based on Last.fm play counts.
 * Creates a playlist where for each play count level (from 0 to max),
 * a shuffled set of all songs with play counts up to that level is added.
 *
 * For example, if max play count is 3:
 * - First, a shuffled list of all songs with 0 or 1 play is added.
 * - Second, a shuffled list of all songs with 0, 1, or 2 plays is added.
 * - Third, a shuffled list of all songs with 0, 1, 2, or 3 plays is added.
 *
 * @param {TrackWithPlayCount[]} tracks - An array of track objects, now with 'lastFmPlayCount' property.
 * @returns {TrackWithPlayCount[]} The shuffled and cumulatively ordered array of track objects.
 */
function applyShuffleLogic(tracks: TrackWithPlayCount[]): TrackWithPlayCount[] {
    if (tracks.length === 0) {
        return [];
    }

    // Group tracks by their effective play count (0 for null/undefined)
    const playCountGroups: Map<number, TrackWithPlayCount[]> = new Map(); // Map<number, Array<Track>>
    let maxPlayCount: number = 0;

    for (const track of tracks) {
        const effectivePlayCount: number = track.lastFmPlayCount !== null ? track.lastFmPlayCount : 0;
        if (!playCountGroups.has(effectivePlayCount)) {
            playCountGroups.set(effectivePlayCount, []);
        }
        playCountGroups.get(effectivePlayCount)!.push(track); // Use ! for non-null assertion
        if (effectivePlayCount > maxPlayCount) {
            maxPlayCount = effectivePlayCount;
        }
    }

    const finalPlaylist: TrackWithPlayCount[] = [];
    let cumulativePool: TrackWithPlayCount[] = []; // This will hold all songs with play count <= current iteration 'i'

    // Iterate through play count levels from 0 up to maxPlayCount
    for (let i = 0; i <= maxPlayCount; i++) {
        // Add songs from the current playCount group to the cumulative pool
        if (playCountGroups.has(i)) {
            cumulativePool.push(...playCountGroups.get(i)!); // Use ! for non-null assertion
        }

        // If the cumulative pool has songs, shuffle it and add to the final playlist
        if (cumulativePool.length > 0) {
            const shuffledCurrentPool: TrackWithPlayCount[] = [...cumulativePool]; // Create a copy to shuffle
            // Fisher-Yates shuffle
            for (let j = shuffledCurrentPool.length - 1; j > 0; j--) {
                const k: number = Math.floor(Math.random() * (j + 1));
                [shuffledCurrentPool[j], shuffledCurrentPool[k]] = [shuffledCurrentPool[k], shuffledCurrentPool[j]];
            }
            finalPlaylist.push(...shuffledCurrentPool);
        }
    }

    return finalPlaylist;
}

/**
 * Creates a new playlist for the user.
 * @param {string} accessToken - The Spotify access token.
 * @param {string} userId - The Spotify user ID.
 * @param {string} playlistName - The name of the new playlist.
 * @returns {Promise<any>} The created playlist object.
 */
async function createPlaylist(accessToken: string, userId: string, playlistName: string): Promise<any> {
    const playlistData: Record<string, string | boolean> = {
        name: playlistName,
        public: false, // Make it private by default
        collaborative: false,
        description: 'Shuffled Liked Songs (cumulative by play count) - generated by Cloudflare Worker'
    };

    return spotifyApiFetch(`${SPOTIFY_API_BASE_URL}/users/${userId}/playlists`, accessToken, {
        method: 'POST',
        body: JSON.stringify(playlistData),
    });
}

/**
 * Adds tracks to a playlist. Spotify API limits to 100 tracks per request.
 * @param {string} accessToken - The Spotify access token.
 * @param {string} playlistId - The ID of the playlist.
 * @param {string[]} trackUris - An array of Spotify track URIs (e.g., ['spotify:track:ID1', 'spotify:track:ID2']).
 * @returns {Promise<void>}
 */
async function addTracksToPlaylist(accessToken: string, playlistId: string, trackUris: string[]): Promise<void> {
    const batchSize: number = 100;
    for (let i = 0; i < trackUris.length; i += batchSize) {
        const batch: string[] = trackUris.slice(i, i + batchSize);
        await spotifyApiFetch(`${SPOTIFY_API_BASE_URL}/playlists/${playlistId}/tracks`, accessToken, {
            method: 'POST',
            body: JSON.stringify({ uris: batch }),
        });
        console.log(`Added batch of ${batch.length} tracks to playlist ${playlistId}`);
    }
}

// --- Main Worker Event Listener ---

addEventListener('fetch', (event: FetchEvent) => {
    event.respondWith(handleRequest(event.request));
});

async function handleRequest(request: Request): Promise<Response> {
    const url: URL = new URL(request.url);
    const path: string = url.pathname;
    const userId: string = 'default_user'; // A simple fixed user ID for single-user setup. For multi-user, derive from auth.

    try {
        if (path === '/login') {
            // Step 1: Redirect to Spotify authorization page
            const state: string = generateRandomString(16);
            const scope: string = 'user-library-read playlist-modify-private playlist-modify-public user-read-private'; // Request necessary scopes

            const queryParams: string = encodeQueryParams({
                response_type: 'code',
                client_id: SPOTIFY_CLIENT_ID,
                scope: scope,
                redirect_uri: SPOTIFY_REDIRECT_URI,
                state: state,
            });

            return Response.redirect(`${SPOTIFY_AUTH_URL}?${queryParams}`, 302);

        } else if (path === '/callback') {
            // Step 2: Handle Spotify callback and exchange code for tokens
            const code: string | null = url.searchParams.get('code');
            const state: string | null = url.searchParams.get('state');
            const error: string | null = url.searchParams.get('error');

            if (error) {
                return new Response(`Spotify authorization error: ${error}`, { status: 400 });
            }
            if (!code) {
                return new Response('Missing authorization code in callback.', { status: 400 });
            }
            // In a real app, you'd verify the 'state' parameter to prevent CSRF.
            // For this example, we'll skip state verification for simplicity.

            const client_id: string = SPOTIFY_CLIENT_ID;
            const client_secret: string = SPOTIFY_CLIENT_SECRET;
            const redirect_uri: string = SPOTIFY_REDIRECT_URI;

            const authHeader: string = btoa(`${client_id}:${client_secret}`);

            const tokenResponse: Response = await fetch(SPOTIFY_TOKEN_URL, {
                method: 'POST',
                headers: {
                    'Authorization': `Basic ${authHeader}`,
                    'Content-Type': 'application/x-www-form-urlencoded',
                },
                body: encodeQueryParams({
                    grant_type: 'authorization_code',
                    code: code,
                    redirect_uri: redirect_uri,
                }),
            });

            if (!tokenResponse.ok) {
                const errorData: any = await tokenResponse.json().catch(() => ({ message: 'Unknown error' }));
                console.error('Token exchange failed:', errorData);
                return new Response(`Failed to exchange code for token: ${tokenResponse.status} - ${errorData.error_description || errorData.error}`, { status: 500 });
            }

            const tokenData: any = await tokenResponse.json();
            const accessToken: string = tokenData.access_token;
            const refreshToken: string = tokenData.refresh_token;
            const expiresIn: number = tokenData.expires_in; // seconds

            // Store tokens in KV for future use
            await SPOTIFY_TOKENS.put(userId, JSON.stringify({
                access_token: accessToken,
                refresh_token: refreshToken,
                expires_at: Date.now() + (expiresIn * 1000)
            } as SpotifyTokenData));

            return new Response('Successfully logged in to Spotify! You can now go to /shuffle to create a playlist.', { status: 200 });

        } else if (path === '/shuffle') {
            // Step 3: Fetch liked songs, shuffle, and create playlist
            let accessToken: string;
            try {
                accessToken = await getValidAccessToken(userId);
            } catch (e: any) {
                console.error('Error getting access token:', e);
                return new Response(`Authentication required. Please visit /login first. Error: ${e.message}`, { status: 401 });
            }

            // Get current user's profile to get their ID
            const userProfile: any = await spotifyApiFetch(`${SPOTIFY_API_BASE_URL}/me`, accessToken);
            const spotifyUserId: string = userProfile.id;

            // Fetch all liked songs from Spotify
            const likedTracks: SpotifyTrack[] = await fetchAllLikedSongs(accessToken);
            if (likedTracks.length === 0) {
                return new Response('No liked songs found in your Spotify library.', { status: 200 });
            }

            // Fetch Last.fm play counts for each track
            console.log(`Fetching Last.fm play counts for ${likedTracks.length} tracks...`);
            const tracksWithPlayCounts: TrackWithPlayCount[] = [];
            for (const track of likedTracks) {
                const artistName: string = track.artists.map((a: { name: string }) => a.name).join(', ');
                const trackName: string = track.name;
                const playCount: number | null = await fetchLastFmPlayCount(artistName, trackName, LASTFM_API_KEY, LASTFM_USERNAME);
                tracksWithPlayCounts.push({ ...track, lastFmPlayCount: playCount });
            }
            console.log('Finished fetching Last.fm play counts.');

            // Apply custom shuffle logic based on Last.fm play counts (cumulative)
            const shuffledTracks: TrackWithPlayCount[] = applyShuffleLogic(tracksWithPlayCounts);

            // Create a new playlist
            const playlistName: string = `Shuffled Liked Songs (Cumulative Plays) - ${new Date().toLocaleString()}`;
            const newPlaylist: any = await createPlaylist(accessToken, spotifyUserId, playlistName);
            const playlistId: string = newPlaylist.id;
            const playlistUrl: string = newPlaylist.external_urls.spotify;

            // Add shuffled tracks to the new playlist
            const trackUris: string[] = shuffledTracks.map((track: TrackWithPlayCount) => track.uri);
            await addTracksToPlaylist(accessToken, playlistId, trackUris);

            return new Response(`Successfully created and populated playlist: <a href="${playlistUrl}" target="_blank">${playlistName}</a> with ${shuffledTracks.length} songs.`, {
                headers: { 'Content-Type': 'text/html' },
                status: 200
            });

        } else {
            // Default route
            return new Response(`
                <h1>Spotify Liked Songs Shuffler Worker</h1>
                <p>Welcome! This worker helps you shuffle your Spotify liked songs into a new playlist, prioritizing least played songs using Last.fm data.</p>
                <p><strong>Important:</strong> You need to configure Spotify and Last.fm API keys and your Last.fm username as environment variables in your Cloudflare Worker settings.</p>
                <ul>
                    <li><a href="/login">Login with Spotify</a>: Authorize this worker to access your Spotify account.</li>
                    <li><a href="/shuffle">Shuffle Liked Songs</a>: After logging in, visit this URL to fetch your liked songs, retrieve Last.fm play counts, apply cumulative shuffle logic, and create a new Spotify playlist.</li>
                </ul>
                <p>Make sure you have set up the required environment variables (SPOTIFY_CLIENT_ID, SPOTIFY_CLIENT_SECRET, SPOTIFY_REDIRECT_URI, <strong>LASTFM_API_KEY, LASTFM_USERNAME</strong>) and a KV namespace (SPOTIFY_TOKENS) in your Cloudflare Worker settings.</p>
            `, { headers: { 'Content-Type': 'text/html' } });
        }
    } catch (error: any) {
        console.error('Worker error:', error);
        return new Response(`An error occurred: ${error.message}`, { status: 500 });
    }
}
