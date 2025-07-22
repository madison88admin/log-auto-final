// src/mocks/handlers.js
import { http, HttpResponse } from 'msw';

export const handlers = [
  http.post('http://localhost:8000/generate-reports/', async ({ request }) => {
    // Create a fake Excel file as a Blob
    const fakeExcel = new Blob(['Fake Excel content'], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    // Return a mocked response using HttpResponse
    return new HttpResponse(fakeExcel, {
      status: 200,
      headers: { 'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }
    });
  }),
];