/**
 * Cloudflare Worker for Spotify Liked Songs Shuffle with Last.fm Play Counts (TypeScript)
 *
 * This worker authenticates with Spotify, fetches a user's liked songs,
 * then fetches the user's top tracks from Last.fm to get play counts.
 * It matches Spotify songs to Last.fm top tracks, assigns play counts,
 * and creates a new Spotify playlist where songs are cumulatively shuffled
 * based on their play count (least played first).
 *
 * ONLY SONGS THAT HAVE A PLAY COUNT RECORDED IN LAST.FM ARE INCLUDED.
 *
 * API REQUESTS ARE NOW LIMITED ONLY WHEN ADDING TRACKS TO THE SPOTIFY PLAYLIST.
 * This means fetching all liked songs and Last.fm top tracks will attempt to get
 * the full available data, but the final playlist size will be capped by the
 * request limit for adding tracks.
 *
 * The /shuffle endpoint returns immediately, and the playlist creation
 * happens asynchronously. A new /status/:processId endpoint allows checking
 * the progress and results.
 *
 * Environment Variables Required:
 * - SPOTIFY_CLIENT_ID: Your Spotify Application Client ID
 * - SPOTIFY_CLIENT_SECRET: Your Spotify Application Client Secret
 * - SPOTIFY_REDIRECT_URI: The redirect URI configured in your Spotify App (e.g., https://your-worker-domain.workers.dev/callback)
 * - LASTFM_API_KEY: Your Last.fm API Key
 * - LASTFM_USERNAME: Your Last.fm Username (for fetching your specific play counts)
 *
 * KV Namespace Required:
 * - SPOTIFY_TOKENS: A KV namespace to store user-specific Spotify access and refresh tokens,
 * and also the status of playlist generation processes.
 */

// Declare global types for Cloudflare Worker environment variables and KV namespace
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

// KV key prefix for storing playlist generation status
const STATUS_KEY_PREFIX: string = 'playlist_status_';

// --- API Request Limits ---
// These limits are designed to keep the total subrequests under 50.
// Fetching liked songs and Last.fm top tracks will now attempt to get all data.
// The primary limit is now on adding tracks to the playlist.
const SPOTIFY_PLAYLIST_ADD_MAX_REQUESTS: number = 40;     // Max 40 Spotify API calls for adding tracks (40 * 100 = 4000 tracks)
                                                         // This leaves ~10 requests for initial auth, fetching liked songs,
                                                         // and fetching Last.fm top tracks.

// Configuration for Last.fm API fetching (for top tracks, not individual lookups)
const LASTFM_TOP_TRACKS_LIMIT_PER_PAGE: number = 1000; // Max limit per page for user.getTopTracks
const LASTFM_PAGE_DELAY_MS: number = 200; // Delay between fetching Last.fm top track pages

// --- Type Definitions ---

interface SpotifyTrack {
    id: string;
    name: string;
    artists: Array<{ name: string }>;
    uri: string;
    [key: string]: any; // Allow for other properties not explicitly defined
}

interface TrackWithPlayCount extends SpotifyTrack {
    lastFmPlayCount: number | null; // Can be null if not found in Last.fm
}

interface SpotifyTokenData {
    access_token: string;
    refresh_token: string;
    expires_at: number; // Unix timestamp in milliseconds
}

interface PlaylistStatus {
    status: 'pending' | 'fetching_liked_songs' | 'fetching_lastfm_top_tracks' | 'matching_tracks' | 'applying_shuffle_logic' | 'creating_playlist' | 'adding_tracks' | 'completed' | 'failed';
    message: string;
    playlistUrl?: string;
    error?: string;
    timestamp: number;
    progress?: string; // Added for progress updates
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
            ...(options.headers as Record<string, string>),
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
    const newRefreshToken: string = data.refresh_token || refreshToken;

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
 * Handles pagination until all songs are fetched or an error occurs.
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
 * Fetches a user's top tracks from Last.fm, handling pagination.
 * Fetches all available pages up to Last.fm's internal limits.
 * @param {string} lastFmApiKey - Your Last.fm API key.
 * @param {string} lastFmUsername - Your Last.fm username.
 * @returns {Promise<Array<{ artist: string, name: string, playcount: number }>>} An array of top track objects.
 */
async function fetchLastFmTopTracks(lastFmApiKey: string, lastFmUsername: string): Promise<Array<{ artist: string, name: string, playcount: number }>> {
    let allTopTracks: Array<{ artist: string, name: string, playcount: number }> = [];
    let page = 1;
    let totalPages = 1; // Initialize to 1 to enter the loop

    while (page <= totalPages) {
        const params: Record<string, string | number> = {
            method: 'user.getTopTracks',
            user: lastFmUsername,
            api_key: lastFmApiKey,
            limit: LASTFM_TOP_TRACKS_LIMIT_PER_PAGE,
            page: page,
            format: 'json'
        };
        const url: string = `${LASTFM_API_BASE_URL}?${encodeQueryParams(params)}`;

        try {
            const response: Response = await fetch(url);

            if (!response.ok) {
                console.error(`Last.fm API Error (${response.status}) fetching top tracks for page ${page}`);
                break;
            }
            const data: any = await response.json();

            if (data && data.toptracks && data.toptracks.track) {
                allTopTracks = allTopTracks.concat(
                    data.toptracks.track.map((t: any) => ({
                        artist: t.artist.name,
                        name: t.name,
                        playcount: parseInt(t.playcount, 10)
                    }))
                );
                totalPages = parseInt(data.toptracks['@attr'].totalPages, 10);
                page++;

                // Add delay between pages to respect Last.fm rate limits
                if (page <= totalPages) {
                    await new Promise(resolve => setTimeout(resolve, LASTFM_PAGE_DELAY_MS));
                }
            } else {
                console.warn('Last.fm top tracks response missing expected data structure.');
                break;
            }
        } catch (error: any) {
            console.error(`Error fetching Last.fm top tracks for page ${page}:`, error);
            break; // Break loop on error
        }
    }
    return allTopTracks;
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
 * @returns {Promise<TrackWithPlayCount[]>} The shuffled and cumulatively ordered array of track objects.
 */
async function applyShuffleLogic(tracks: TrackWithPlayCount[]): Promise<TrackWithPlayCount[]> {
    if (tracks.length === 0) {
        return [];
    }

    // Group tracks by their effective play count (0 for null/undefined)
    const playCountGroups: Map<number, TrackWithPlayCount[]> = new Map(); // Map<number, Array<Track>>
    let maxPlayCount: number = 0;

    for (const track of tracks) {
        // Ensure that tracks with null play counts (not found in Last.fm) are not processed here.
        // This function should only receive tracks that already have a play count (even if 0).
        const effectivePlayCount: number = track.lastFmPlayCount !== null ? track.lastFmPlayCount : 0; // Should not be null at this point
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
 * Adds tracks to a playlist, respecting a maximum number of API requests.
 * Spotify API limits to 100 tracks per request.
 * @param {string} accessToken - The Spotify access token.
 * @param {string} playlistId - The ID of the playlist.
 * @param {string[]} trackUris - An array of Spotify track URIs (e.g., ['spotify:track:ID1', 'spotify:track:ID2']).
 * @param {number} maxRequests - The maximum number of API requests to make for adding tracks.
 * @returns {Promise<void>}
 */
async function addTracksToPlaylist(accessToken: string, playlistId: string, trackUris: string[], maxRequests: number): Promise<void> {
    const batchSize: number = 100;
    let requestCount = 0;

    for (let i = 0; i < trackUris.length && requestCount < maxRequests; i += batchSize) {
        const batch: string[] = trackUris.slice(i, i + batchSize);
        await spotifyApiFetch(`${SPOTIFY_API_BASE_URL}/playlists/${playlistId}/tracks`, accessToken, {
            method: 'POST',
            body: JSON.stringify({ uris: batch }),
        });
        requestCount++;
        console.log(`Added batch of ${batch.length} tracks to playlist ${playlistId} (Request ${requestCount}/${maxRequests})`);
    }
    if (requestCount >= maxRequests && trackUris.length > requestCount * batchSize) {
        console.warn(`Stopped adding tracks to playlist after ${maxRequests} requests due to limit.`);
    }
}

/**
 * Initiates the asynchronous playlist generation process.
 * This function is called via event.waitUntil() and updates KV with status.
 * @param {string} userId - The user ID.
 * @param {string} accessToken - The Spotify access token.
 * @param {string} spotifyUserId - The Spotify user ID.
 * @param {string} processId - The unique ID for this process.
 */
async function startPlaylistGeneration(userId: string, accessToken: string, spotifyUserId: string, processId: string): Promise<void> {
    const statusKey = STATUS_KEY_PREFIX + processId;

    try {
        await SPOTIFY_TOKENS.put(statusKey, JSON.stringify({
            status: 'fetching_liked_songs',
            message: `Fetching all liked songs from Spotify...`, // Removed max limit message
            timestamp: Date.now(),
            progress: '0%'
        } as PlaylistStatus));

        // Fetch all liked songs (no request limit applied here)
        const likedTracks: SpotifyTrack[] = await fetchAllLikedSongs(accessToken);
        if (likedTracks.length === 0) {
            await SPOTIFY_TOKENS.put(statusKey, JSON.stringify({
                status: 'completed',
                message: 'No liked songs found in your Spotify library.',
                timestamp: Date.now(),
                progress: '100%'
            } as PlaylistStatus));
            return;
        }

        await SPOTIFY_TOKENS.put(statusKey, JSON.stringify({
            status: 'fetching_lastfm_top_tracks',
            message: `Fetching all Last.fm top tracks (this may take a while for large libraries)...`, // Removed max limit message
            timestamp: Date.now(),
            progress: '10%'
        } as PlaylistStatus));

        // Fetch all Last.fm top tracks (no request limit applied here)
        const lastFmTopTracks = await fetchLastFmTopTracks(LASTFM_API_KEY, LASTFM_USERNAME);
        const lastFmPlayCountMap = new Map<string, number>(); // Key: "ArtistName - TrackName" (normalized)

        for (const track of lastFmTopTracks) {
            const normalizedArtist = track.artist.toLowerCase().trim();
            const normalizedTrack = track.name.toLowerCase().trim();
            lastFmPlayCountMap.set(`${normalizedArtist} - ${normalizedTrack}`, track.playcount);
        }

        await SPOTIFY_TOKENS.put(statusKey, JSON.stringify({
            status: 'matching_tracks',
            message: `Matching ${likedTracks.length} Spotify songs with Last.fm data and filtering...`,
            timestamp: Date.now(),
            progress: '50%'
        } as PlaylistStatus));

        const tracksWithPlayCounts: TrackWithPlayCount[] = [];
        for (const track of likedTracks) {
            const normalizedArtist = track.artists.map(a => a.name).join(', ').toLowerCase().trim();
            const normalizedTrack = track.name.toLowerCase().trim();
            const playCount = lastFmPlayCountMap.get(`${normalizedArtist} - ${normalizedTrack}`);

            // Only push the track if a play count was found in Last.fm
            if (playCount !== undefined) {
                tracksWithPlayCounts.push({ ...track, lastFmPlayCount: playCount });
            }
        }
        console.log(`Finished matching Last.fm play counts. ${tracksWithPlayCounts.length} tracks included.`);

        if (tracksWithPlayCounts.length === 0) {
            await SPOTIFY_TOKENS.put(statusKey, JSON.stringify({
                status: 'completed',
                message: 'No liked songs found with Last.fm play counts. Playlist not created.',
                timestamp: Date.now(),
                progress: '100%'
            } as PlaylistStatus));
            return;
        }


        await SPOTIFY_TOKENS.put(statusKey, JSON.stringify({
            status: 'applying_shuffle_logic',
            message: 'Applying cumulative shuffle logic...',
            timestamp: Date.now(),
            progress: '90%'
        } as PlaylistStatus));

        const shuffledTracks: TrackWithPlayCount[] = await applyShuffleLogic(tracksWithPlayCounts);

        await SPOTIFY_TOKENS.put(statusKey, JSON.stringify({
            status: 'creating_playlist',
            message: 'Creating new Spotify playlist...',
            timestamp: Date.now(),
            progress: '95%'
        } as PlaylistStatus));

        const playlistName: string = `Shuffled Liked Songs (Cumulative Plays) - ${new Date().toLocaleString()}`;
        const newPlaylist: any = await createPlaylist(accessToken, spotifyUserId, playlistName);
        const playlistId: string = newPlaylist.id;
        const playlistUrl: string = newPlaylist.external_urls.spotify;

        await SPOTIFY_TOKENS.put(statusKey, JSON.stringify({
            status: 'adding_tracks',
            message: `Adding ${shuffledTracks.length} tracks to playlist (limited to ${SPOTIFY_PLAYLIST_ADD_MAX_REQUESTS * 100} tracks)...`, // Updated message
            timestamp: Date.now(),
            progress: '98%'
        } as PlaylistStatus));

        const trackUris: string[] = shuffledTracks.map((track: TrackWithPlayCount) => track.uri);
        // Apply limit only here
        await addTracksToPlaylist(accessToken, playlistId, trackUris, SPOTIFY_PLAYLIST_ADD_MAX_REQUESTS);

        await SPOTIFY_TOKENS.put(statusKey, JSON.stringify({
            status: 'completed',
            message: `Successfully created and populated playlist: ${playlistName}`,
            playlistUrl: playlistUrl,
            timestamp: Date.now(),
            progress: '100%'
        } as PlaylistStatus), { expirationTtl: 3600 }); // Store completed status for 1 hour

    } catch (error: any) {
        console.error(`Playlist generation failed for process ${processId}:`, error);
        await SPOTIFY_TOKENS.put(statusKey, JSON.stringify({
            status: 'failed',
            message: 'Playlist generation failed.',
            error: error.message || 'Unknown error',
            timestamp: Date.now(),
            progress: '100%'
        } as PlaylistStatus), { expirationTtl: 3600 }); // Store failed status for 1 hour
    }
}

// --- Main Worker Event Listener ---

addEventListener('fetch', (event: FetchEvent) => {
    event.respondWith(handleRequest(event.request, event));
});

async function handleRequest(request: Request, event: FetchEvent): Promise<Response> {
    const url: URL = new URL(request.url);
    const path: string = url.pathname;
    const userId: string = 'default_user'; // A simple fixed user ID for single-user setup. For multi-user, derive from auth.

    try {
        if (path === '/login') {
            const state: string = generateRandomString(16);
            const scope: string = 'user-library-read playlist-modify-private playlist-modify-public user-read-private';

            const queryParams: string = encodeQueryParams({
                response_type: 'code',
                client_id: SPOTIFY_CLIENT_ID,
                scope: scope,
                redirect_uri: SPOTIFY_REDIRECT_URI,
                state: state,
            });

            return Response.redirect(`${SPOTIFY_AUTH_URL}?${queryParams}`, 302);

        } else if (path === '/callback') {
            const code: string | null = url.searchParams.get('code');
            const state: string | null = url.searchParams.get('state');
            const error: string | null = url.searchParams.get('error');

            if (error) {
                return new Response(`Spotify authorization error: ${error}`, { status: 400 });
            }
            if (!code) {
                return new Response('Missing authorization code in callback.', { status: 400 });
            }

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
            const expiresIn: number = tokenData.expires_in;

            await SPOTIFY_TOKENS.put(userId, JSON.stringify({
                access_token: accessToken,
                refresh_token: refreshToken,
                expires_at: Date.now() + (expiresIn * 1000)
            } as SpotifyTokenData));

            return new Response('Successfully logged in to Spotify! You can now go to /shuffle to create a playlist.', { status: 200 });

        } else if (path === '/shuffle') {
            let accessToken: string;
            try {
                accessToken = await getValidAccessToken(userId);
            } catch (e: any) {
                console.error('Error getting access token:', e);
                return new Response(`Authentication required. Please visit /login first. Error: ${e.message}`, { status: 401 });
            }

            const userProfile: any = await spotifyApiFetch(`${SPOTIFY_API_BASE_URL}/me`, accessToken);
            const spotifyUserId: string = userProfile.id;

            const processId: string = generateRandomString(20); // Unique ID for this process
            const statusKey = STATUS_KEY_PREFIX + processId;

            // Store initial pending status
            await SPOTIFY_TOKENS.put(statusKey, JSON.stringify({
                status: 'pending',
                message: 'Playlist generation process initiated.',
                timestamp: Date.now(),
                progress: '0%'
            } as PlaylistStatus));

            // Start the playlist generation in the background
            event.waitUntil(startPlaylistGeneration(userId, accessToken, spotifyUserId, processId));

            return new Response(`Playlist generation started! Check status at <a href="/status/${processId}" target="_blank">/status/${processId}</a>`, {
                headers: { 'Content-Type': 'text/html' },
                status: 202
            });

        } else if (path.startsWith('/status/')) {
            const processId: string = path.substring('/status/'.length);
            const statusKey = STATUS_KEY_PREFIX + processId;

            const statusDataStr: string | null = await SPOTIFY_TOKENS.get(statusKey);

            if (!statusDataStr) {
                return new Response('Process ID not found or expired. Status messages are retained for 1 hour after completion/failure.', { status: 404 });
            }

            const status: PlaylistStatus = JSON.parse(statusDataStr);
            return new Response(JSON.stringify(status, null, 2), { // Pretty print JSON
                headers: { 'Content-Type': 'application/json' },
                status: 200
            });

        } else {
            return new Response(`
                <h1>Spotify Liked Songs Shuffler Worker</h1>
                <p>Welcome! This worker helps you shuffle your Spotify liked songs into a new playlist, prioritizing least played songs using Last.fm data.</p>
                <p><strong>Important:</strong> You need to configure Spotify and Last.fm API keys and your Last.fm username as environment variables in your Cloudflare Worker settings.</p>
                <ul>
                    <li><a href="/login">Login with Spotify</a>: Authorize this worker to access your Spotify account.</li>
                    <li><a href="/shuffle">Shuffle Liked Songs (Async)</a>: Visit this URL to initiate the playlist generation. You'll get a process ID to check its status.</li>
                    <li><strong>Check Status:</strong> After starting a shuffle, append the process ID to <code>/status/</code> (e.g., <code>/status/YOUR_PROCESS_ID</code>) to get updates.</li>
                </ul>
                <p>Make sure you have set up the required environment variables (SPOTIFY_CLIENT_ID, SPOTIFY_CLIENT_SECRET, SPOTIFY_REDIRECT_URI, <strong>LASTFM_API_KEY, LASTFM_USERNAME</strong>) and a KV namespace (SPOTIFY_TOKENS) in your Cloudflare Worker settings.</p>
            `, { headers: { 'Content-Type': 'text/html' } });
        }
    } catch (error: any) {
        console.error('Worker error:', error);
        return new Response(`An error occurred: ${error.message}`, { status: 500 });
    }
}
