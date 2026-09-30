# Si Vercel muestra FUNCTION_INVOCATION_FAILED

1. Project → Settings → Environment Variables
   - Debe existir DATABASE_URL (toda la connection string de Neon)
   - Environments: Production + Preview

2. Project → Deployments → los 3 puntitos del último deploy → Redeploy

3. Logs: Deployments → clic en el deploy → Runtime Logs / Building
   Copia el error rojo si sigue fallando.
