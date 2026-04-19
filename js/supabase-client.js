// Supabase configuration
const SUPABASE_URL = 'https://wocrvonmsvpfisflpjez.supabase.co';
const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6IndvY3J2b25tc3ZwZmlzZmxwamV6Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzY1NzgwMTEsImV4cCI6MjA5MjE1NDAxMX0.SE3vtlD2UenX4cbgmGhtNUfaeIgWuMXXJnaYZPz2K7w';

const { createClient } = supabase;
const sb = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);
