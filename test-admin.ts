import { initializeApp } from 'firebase-admin/app';
import { getAuth } from 'firebase-admin/auth';

async function test() {
  try {
    const app = initializeApp();
    console.log("Admin app initialized");
    const list = await getAuth(app).listUsers(1);
    console.log("Successfully listed users:", list.users.length);
  } catch (e) {
    console.error("Error:", e);
  }
}
test();
