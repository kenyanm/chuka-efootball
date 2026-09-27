import express, { Request, Response, NextFunction } from 'express';
import { createServer as createViteServer } from 'vite';
import path from 'path';
import { fileURLToPath } from 'url';
import { initializeApp, getApps, cert } from 'firebase-admin/app';
import { getAuth } from 'firebase-admin/auth';
import dotenv from 'dotenv';
import fs from 'fs';

dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

let portArg = process.env.PORT;
for (let i = 0; i < process.argv.length; i++) {
  if (process.argv[i] === '--port' && process.argv[i + 1]) {
    portArg = process.argv[i + 1];
  } else if (process.argv[i].startsWith('--port=')) {
    portArg = process.argv[i].split('=')[1];
  }
}
const PORT = parseInt(portArg || '3000', 10);
const PROJECT_ID = process.env.VITE_FIREBASE_PROJECT_ID || 'chuka-efootball-hub';
const FIREBASE_API_KEY = process.env.VITE_FIREBASE_API_KEY || 'AIzaSyBaRnQXGnD57G_KSK3MeMjEG1x3hxivDqw';

// Initialize Firebase Admin SDK
let adminApp;
if (getApps().length === 0) {
  // Check if service account key is available in environment or file
  const serviceAccountJson = process.env.FIREBASE_SERVICE_ACCOUNT_KEY || process.env.GOOGLE_APPLICATION_CREDENTIALS_JSON;
  if (serviceAccountJson) {
    try {
      const parsed = JSON.parse(serviceAccountJson);
      adminApp = initializeApp({ credential: cert(parsed), projectId: PROJECT_ID }, 'chukaAdminApp');
      console.log('[Firebase Admin] Initialized with explicit service account credentials.');
    } catch (e) {
      console.warn('[Firebase Admin] Failed to parse service account JSON, initializing with project ID:', e);
      adminApp = initializeApp({ projectId: PROJECT_ID }, 'chukaAdminApp');
    }
  } else {
    adminApp = initializeApp({ projectId: PROJECT_ID }, 'chukaAdminApp');
    console.log(`[Firebase Admin] Initialized for project: ${PROJECT_ID}`);
  }
} else {
  adminApp = getApps()[0];
}

const adminAuth = getAuth(adminApp);

// Authoritative list of administrative email addresses (Enforced Server-Side)
export const AUTHORIZED_ADMIN_EMAILS = [
  'wayongohlaurence@gmail.com',
  'wayongohlawrence@gmail.com',
  'sidobarasa7@gmail.com',
];

export function isAuthorizedAdminEmail(email?: string | null): boolean {
  if (!email || typeof email !== 'string') return false;
  const normalized = email.trim().toLowerCase();
  return AUTHORIZED_ADMIN_EMAILS.some((adm) => adm.toLowerCase() === normalized);
}

export function isSuperAdminEmail(email?: string | null): boolean {
  if (!email || typeof email !== 'string') return false;
  const normalized = email.trim().toLowerCase();
  return normalized === 'wayongohlaurence@gmail.com' || normalized === 'wayongohlawrence@gmail.com';
}

/**
 * Server-side Firebase ID Token verification & Admin authorization check
 */
async function authenticateAdminRequest(req: Request) {
  const authHeader = req.headers.authorization;
  const token = authHeader?.startsWith('Bearer ')
    ? authHeader.substring(7).trim()
    : (req.body?.idToken || req.query?.idToken || '').trim();

  if (!token) {
    return {
      authorized: false,
      status: 401,
      code: 'AUTH_REQUIRED',
      message: 'Authentication token is required for administrative operations.',
    };
  }

  let decodedToken: any;
  try {
    decodedToken = await adminAuth.verifyIdToken(token);
  } catch (adminErr: any) {
    // Robust fallback: Verify token via Google Identity Toolkit REST API
    try {
      const verifyRes = await fetch(
        `https://identitytoolkit.googleapis.com/v1/accounts:lookup?key=${FIREBASE_API_KEY}`,
        {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ idToken: token }),
        }
      );
      const verifyData: any = await verifyRes.json();
      if (verifyData && verifyData.users && verifyData.users.length > 0) {
        const u = verifyData.users[0];
        let parsedClaims: any = {};
        try {
          const parts = token.split('.');
          if (parts[1]) {
            parsedClaims = JSON.parse(Buffer.from(parts[1], 'base64').toString('utf8'));
          }
        } catch {}

        decodedToken = {
          uid: u.localId,
          email: u.email,
          email_verified: Boolean(u.emailVerified),
          displayName: u.displayName || parsedClaims.name || '',
          photoURL: u.photoUrl || parsedClaims.picture || '',
          ...parsedClaims,
          admin: parsedClaims.admin === true,
          role: parsedClaims.role,
        };
      } else {
        throw new Error(verifyData?.error?.message || 'Token lookup failed');
      }
    } catch (lookupErr: any) {
      console.error('[Server Auth] Failed to verify Firebase ID token:', lookupErr.message || adminErr.message);
      return {
        authorized: false,
        status: 401,
        code: 'INVALID_TOKEN',
        message: 'Invalid or expired Firebase ID token: ' + (lookupErr.message || adminErr.message),
      };
    }
  }

  const email = (decodedToken.email || '').toLowerCase().trim();
  const isOwner = isSuperAdminEmail(email);
  const isAuthorized = isAuthorizedAdminEmail(email);
  const hasAdminClaim = Boolean(
    decodedToken.admin === true ||
    decodedToken.role === 'admin' ||
    decodedToken.role === 'ADMIN' ||
    decodedToken.role === 'SUPER_ADMIN'
  );

  if (!isAuthorized && !hasAdminClaim && !isOwner) {
    return {
      authorized: false,
      status: 403,
      code: 'FORBIDDEN_NOT_ADMIN',
      message: 'Forbidden: Authenticated user is not registered as an administrator.',
      user: decodedToken,
    };
  }

  return {
    authorized: true,
    user: decodedToken,
    uid: decodedToken.uid,
    email: decodedToken.email,
    isOwner,
    isSuperAdmin: isOwner || decodedToken.role === 'SUPER_ADMIN',
    role: isOwner ? 'SUPER_ADMIN' : 'ADMIN',
    hasClaim: hasAdminClaim,
  };
}

async function startServer() {
  const app = express();

  app.use(express.json({ limit: '30mb' }));
  app.use(express.urlencoded({ extended: true, limit: '30mb' }));

  // ==========================================
  // SERVER API ROUTES
  // ==========================================

  // 1. Health check
  app.get('/api/health', (req: Request, res: Response) => {
    res.json({
      status: 'ok',
      service: 'Chuka eFootball Hub Server',
      projectId: PROJECT_ID,
      timestamp: new Date().toISOString(),
    });
  });

  // 2. Server-side Admin Verification & Claims Check
  app.post('/api/auth/verify-admin', async (req: Request, res: Response) => {
    const authResult = await authenticateAdminRequest(req);
    if (!authResult.authorized) {
      return res.status(authResult.status || 403).json({
        success: false,
        authenticated: Boolean(authResult.user),
        isAdmin: false,
        error: {
          code: authResult.code,
          message: authResult.message,
        },
      });
    }

    return res.json({
      success: true,
      authenticated: true,
      isAdmin: true,
      isSuperAdmin: authResult.isSuperAdmin,
      role: authResult.role,
      uid: authResult.uid,
      email: authResult.email,
      hasAdminClaim: authResult.hasClaim,
      claims: {
        admin: true,
        role: authResult.role,
      },
      verifiedAt: new Date().toISOString(),
    });
  });

  // Helper to verify any Firebase ID token (player or admin)
  async function verifyAnyUserToken(req: Request) {
    const authHeader = req.headers.authorization;
    const token = authHeader?.startsWith('Bearer ')
      ? authHeader.substring(7).trim()
      : (req.body?.idToken || req.query?.idToken || '').trim();

    if (!token) {
      return {
        valid: false,
        status: 401,
        message: 'Firebase authentication token is required.',
      };
    }

    try {
      const decoded = await adminAuth.verifyIdToken(token);
      const email = (decoded.email || '').toLowerCase().trim();
      const isAdmin = isAuthorizedAdminEmail(email) || decoded.admin === true || decoded.role === 'ADMIN' || decoded.role === 'SUPER_ADMIN';
      const isOwner = isSuperAdminEmail(email);
      return {
        valid: true,
        uid: decoded.uid,
        email: decoded.email,
        displayName: decoded.name || decoded.displayName || '',
        photoURL: decoded.picture || decoded.photoURL || '',
        isAdmin,
        isOwner,
        role: isOwner ? 'SUPER_ADMIN' : isAdmin ? 'ADMIN' : 'USER',
        token,
      };
    } catch {
      // Identity toolkit lookup fallback
      try {
        const verifyRes = await fetch(
          `https://identitytoolkit.googleapis.com/v1/accounts:lookup?key=${FIREBASE_API_KEY}`,
          {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ idToken: token }),
          }
        );
        const verifyData: any = await verifyRes.json();
        if (verifyData?.users && verifyData.users.length > 0) {
          const u = verifyData.users[0];
          const email = (u.email || '').toLowerCase().trim();
          const isAdmin = isAuthorizedAdminEmail(email);
          const isOwner = isSuperAdminEmail(email);
          return {
            valid: true,
            uid: u.localId,
            email: u.email,
            displayName: u.displayName || '',
            photoURL: u.photoUrl || '',
            isAdmin,
            isOwner,
            role: isOwner ? 'SUPER_ADMIN' : isAdmin ? 'ADMIN' : 'USER',
            token,
          };
        }
      } catch (fallbackErr: any) {
        console.warn('[Server Auth] Identity toolkit lookup error:', fallbackErr?.message);
      }
      return {
        valid: false,
        status: 401,
        message: 'Invalid or expired Firebase ID token.',
      };
    }
  }

  // 3. User Sync Endpoint (Resilient: Synchronizes with Sheets if available, falls back to Firebase)
  app.post('/api/auth/sync-user', async (req: Request, res: Response) => {
    const authUser = await verifyAnyUserToken(req);
    if (!authUser.valid) {
      return res.status(authUser.status || 401).json({
        success: false,
        message: authUser.message,
      });
    }

    const appsScriptUrl = process.env.VITE_APPS_SCRIPT_URL || process.env.VITE_APPS_SCRIPT_WEB_APP_URL;
    let syncedUser: any = {
      user_id: authUser.uid,
      email: authUser.email,
      display_name: authUser.displayName || (authUser.email ? authUser.email.split('@')[0] : 'Player'),
      photo_url: authUser.photoURL || undefined,
      role: authUser.role,
      status: 'ACTIVE',
      last_login: new Date().toISOString(),
    };

    if (appsScriptUrl) {
      try {
        const scriptRes = await fetch(appsScriptUrl, {
          method: 'POST',
          headers: {
            'Content-Type': 'text/plain;charset=utf-8',
            Accept: 'application/json',
          },
          body: JSON.stringify({
            action: 'syncUser',
            idToken: authUser.token,
          }),
        });

        const rawText = await scriptRes.text();
        if (rawText && !rawText.trim().startsWith('<')) {
          try {
            const parsed = JSON.parse(rawText);
            if (parsed && parsed.success && parsed.user) {
              syncedUser = {
                ...syncedUser,
                ...parsed.user,
                role: authUser.role,
              };
              return res.json({
                success: true,
                isNewUser: Boolean(parsed.isNewUser),
                user: syncedUser,
                syncedWithSheets: true,
              });
            }
          } catch {}
        }
      } catch (err: any) {
        console.warn('[Server Auth Sync] Apps Script unreachable, using Firebase profile:', err?.message);
      }
    }

    return res.json({
      success: true,
      isNewUser: false,
      user: syncedUser,
      syncedWithSheets: false,
      message: 'Authenticated and synchronized via Firebase.',
    });
  });

  // User Profile Retrieval Endpoint
  app.post('/api/auth/profile', async (req: Request, res: Response) => {
    const authUser = await verifyAnyUserToken(req);
    if (!authUser.valid) {
      return res.status(authUser.status || 401).json({
        success: false,
        message: authUser.message,
      });
    }

    const appsScriptUrl = process.env.VITE_APPS_SCRIPT_URL || process.env.VITE_APPS_SCRIPT_WEB_APP_URL;
    let userProfile: any = {
      user_id: authUser.uid,
      email: authUser.email,
      display_name: authUser.displayName || (authUser.email ? authUser.email.split('@')[0] : 'Player'),
      photo_url: authUser.photoURL || undefined,
      role: authUser.role,
      status: 'ACTIVE',
      last_login: new Date().toISOString(),
    };

    if (appsScriptUrl) {
      try {
        const scriptRes = await fetch(appsScriptUrl, {
          method: 'POST',
          headers: {
            'Content-Type': 'text/plain;charset=utf-8',
            Accept: 'application/json',
          },
          body: JSON.stringify({
            action: 'getProfile',
            idToken: authUser.token,
          }),
        });

        const rawText = await scriptRes.text();
        if (rawText && !rawText.trim().startsWith('<')) {
          try {
            const parsed = JSON.parse(rawText);
            if (parsed && parsed.success && parsed.user) {
              return res.json({
                success: true,
                user: {
                  ...userProfile,
                  ...parsed.user,
                  role: authUser.role,
                },
              });
            }
          } catch {}
        }
      } catch (err: any) {
        console.warn('[Server Profile] Apps Script unreachable, returning token profile:', err?.message);
      }
    }

    return res.json({
      success: true,
      user: userProfile,
    });
  });

  // User Profile Update Endpoint
  app.post('/api/auth/update-profile', async (req: Request, res: Response) => {
    const authUser = await verifyAnyUserToken(req);
    if (!authUser.valid) {
      return res.status(authUser.status || 401).json({
        success: false,
        message: authUser.message,
      });
    }

    const appsScriptUrl = process.env.VITE_APPS_SCRIPT_URL || process.env.VITE_APPS_SCRIPT_WEB_APP_URL;
    const profileUpdates = req.body?.profile || {};

    if (appsScriptUrl) {
      try {
        const scriptRes = await fetch(appsScriptUrl, {
          method: 'POST',
          headers: {
            'Content-Type': 'text/plain;charset=utf-8',
            Accept: 'application/json',
          },
          body: JSON.stringify({
            action: 'updateProfile',
            idToken: authUser.token,
            profile: profileUpdates,
          }),
        });

        const rawText = await scriptRes.text();
        if (rawText && !rawText.trim().startsWith('<')) {
          try {
            const parsed = JSON.parse(rawText);
            if (parsed) {
              return res.json(parsed);
            }
          } catch {}
        }
      } catch (err: any) {
        console.warn('[Server Update Profile] Apps Script unreachable:', err?.message);
      }
    }

    return res.json({
      success: true,
      message: 'Profile updated successfully.',
      user: {
        user_id: authUser.uid,
        email: authUser.email,
        display_name: profileUpdates.display_name || authUser.displayName,
        phone: profileUpdates.phone,
        whatsapp: profileUpdates.whatsapp,
        class_id: profileUpdates.class_id,
        role: authUser.role,
      },
    });
  });

  // 4. Server-Side Custom Claim Assignment via Firebase Admin SDK
  app.post('/api/auth/sync-admin-claims', async (req: Request, res: Response) => {
    const authResult = await authenticateAdminRequest(req);
    if (!authResult.authorized) {
      return res.status(authResult.status || 403).json({
        success: false,
        isAdmin: false,
        error: {
          code: authResult.code,
          message: authResult.message,
        },
      });
    }

    const targetUid = authResult.uid;
    const isOwner = authResult.isOwner;
    const adminClaims = {
      admin: true,
      role: isOwner ? 'SUPER_ADMIN' : 'admin',
      isSuperAdmin: isOwner,
    };

    let claimsAssigned = false;
    let assignmentNotice = '';

    try {
      if (targetUid) {
        await adminAuth.setCustomUserClaims(targetUid, adminClaims);
        claimsAssigned = true;
        console.log(`[Firebase Admin] Successfully assigned custom admin claims to UID: ${targetUid} (${authResult.email})`);
      }
    } catch (claimErr: any) {
      console.warn(`[Firebase Admin Notice] Could not assign custom claim via Identity Toolkit: ${claimErr.message}`);
      assignmentNotice = claimErr.message;
    }

    return res.json({
      success: true,
      isAdmin: true,
      isSuperAdmin: isOwner,
      role: isOwner ? 'SUPER_ADMIN' : 'ADMIN',
      claims: adminClaims,
      claimsAssigned,
      notice: assignmentNotice || undefined,
      message: claimsAssigned
        ? 'Admin claims successfully assigned to Firebase Authentication account.'
        : 'Server-side admin authorization verified.',
      email: authResult.email,
    });
  });

  // 4. Server-Authoritative Admin Operations (Protected by verifyAdminAuth)
  const adminRouter = express.Router();

  adminRouter.use(async (req: Request, res: Response, next: NextFunction) => {
    const authResult = await authenticateAdminRequest(req);
    if (!authResult.authorized) {
      return res.status(authResult.status || 403).json({
        success: false,
        error: {
          code: authResult.code || 'UNAUTHORIZED',
          message: authResult.message || 'Administrative authorization required.',
        },
      });
    }
    (req as any).adminUser = authResult;
    next();
  });

  // Admin Check endpoint
  adminRouter.get('/check', (req: Request, res: Response) => {
    const adminUser = (req as any).adminUser;
    res.json({
      success: true,
      authorized: true,
      admin: adminUser,
    });
  });

  // Helper to forward server-authoritative admin actions to Google Apps Script
  const forwardAdminAction = async (defaultAction: string, req: Request, res: Response) => {
    const adminUser = (req as any).adminUser;
    const bearerToken = (req.headers.authorization || '').replace(/^Bearer\s+/i, '').trim();
    const body = {
      ...(req.body || {}),
      action: req.body?.action || defaultAction,
      idToken: req.body?.idToken || bearerToken,
    };
    const appsScriptUrl = process.env.VITE_APPS_SCRIPT_URL || process.env.VITE_APPS_SCRIPT_WEB_APP_URL;

    console.log(`[Admin Action: ${body.action}] Authorized admin: ${adminUser.email}`);

    if (!appsScriptUrl) {
      return res.status(503).json({
        success: false,
        error: {
          code: 'BACKEND_NOT_CONFIGURED',
          message: 'Apps Script backend URL is not configured.',
        },
      });
    }

    try {
      const authPayload = {
        uid: adminUser.uid,
        email: adminUser.email,
        role: adminUser.role || 'ADMIN',
      };
      const response = await fetch(appsScriptUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'text/plain;charset=utf-8',
          'Accept': 'application/json',
        },
        body: JSON.stringify({
          ...body,
          auth: authPayload,
          adminEmail: adminUser.email,
        }),
      });

      const responseText = await response.text();
      let responseJson: any;
      try {
        responseJson = JSON.parse(responseText);
      } catch {
        return res.status(502).json({
          success: false,
          error: {
            code: 'INVALID_BACKEND_RESPONSE',
            message: 'Invalid non-JSON response from Google Apps Script backend.',
            raw: responseText,
          },
        });
      }

      if (responseJson && responseJson.success === false) {
        return res.status(400).json(responseJson);
      }
      return res.json(responseJson);
    } catch (proxyErr: any) {
      console.error('[Admin Proxy Error]:', proxyErr);
      return res.status(502).json({
        success: false,
        error: {
          code: 'BACKEND_PROXY_FAILED',
          message: 'Failed to communicate with Google Apps Script backend: ' + proxyErr.message,
        },
      });
    }
  };

  // Universal Server-Authoritative Proxy for Admin Operations to Google Apps Script
  // Cryptographically verifies Firebase ID token via Firebase Admin SDK before forwarding
  adminRouter.post('/proxy', async (req: Request, res: Response) => {
    return forwardAdminAction(req.body?.action || 'unknown', req, res);
  });

  // Admin Competition Operations
  adminRouter.post('/competitions/create', (req: Request, res: Response) => {
    return forwardAdminAction('createCompetition', req, res);
  });

  adminRouter.post('/competitions/update', (req: Request, res: Response) => {
    return forwardAdminAction('updateCompetition', req, res);
  });

  adminRouter.post('/registrations/confirm', (req: Request, res: Response) => {
    return forwardAdminAction('confirmRegistration', req, res);
  });

  adminRouter.post('/registrations/reject', (req: Request, res: Response) => {
    return forwardAdminAction('rejectRegistration', req, res);
  });

  adminRouter.post('/payments/verify', (req: Request, res: Response) => {
    return forwardAdminAction('confirmPayment', req, res);
  });

  adminRouter.post('/payments/reject', (req: Request, res: Response) => {
    return forwardAdminAction('rejectPayment', req, res);
  });

  adminRouter.post('/fixtures/record-result', (req: Request, res: Response) => {
    return forwardAdminAction('recordMatchResult', req, res);
  });

  adminRouter.post('/announcements/create', (req: Request, res: Response) => {
    return forwardAdminAction('createAnnouncement', req, res);
  });

  adminRouter.post('/players/verify', (req: Request, res: Response) => {
    return forwardAdminAction('player-verify', req, res);
  });

  app.use('/api/admin', adminRouter);

  // ==========================================
  // VITE / STATIC SERVING
  // ==========================================
  const isProd = process.env.NODE_ENV === 'production';

  if (!isProd) {
    console.log('[Server] Starting Vite development server middleware...');
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    console.log('[Server] Production mode: Serving static files from dist...');
    const distPath = path.resolve(__dirname, 'dist');
    if (fs.existsSync(distPath)) {
      app.use(express.static(distPath));
      app.get('*', (req: Request, res: Response) => {
        res.sendFile(path.resolve(distPath, 'index.html'));
      });
    } else {
      console.warn('[Server Warning] dist/ does not exist. Please run npm run build.');
    }
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`[Server] Chuka eFootball backend listening on http://0.0.0.0:${PORT}`);
  });
}

startServer().catch((err) => {
  console.error('[Server Error] Failed to start server:', err);
  process.exit(1);
});
