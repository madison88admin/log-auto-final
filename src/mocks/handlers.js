// src/mocks/handlers.js
import { rest } from 'msw';

export const handlers = [
  rest.post('http://localhost:8000/generate-reports/', (req, res, ctx) => {
    // Create a fake Excel file as a Blob
    const fakeExcel = new Blob(['Fake Excel content'], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    return res(
      ctx.status(200),
      ctx.body(fakeExcel)
    );
  }),
];