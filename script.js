const canvas = document.getElementById("data-field");
const ctx = canvas.getContext("2d");

const points = [];
let width = 0;
let height = 0;
let animationFrame = 0;

function resize() {
  const ratio = window.devicePixelRatio || 1;
  width = canvas.offsetWidth;
  height = canvas.offsetHeight;
  canvas.width = Math.floor(width * ratio);
  canvas.height = Math.floor(height * ratio);
  ctx.setTransform(ratio, 0, 0, ratio, 0, 0);
  points.length = 0;

  const count = Math.max(42, Math.floor((width * height) / 26000));
  for (let i = 0; i < count; i += 1) {
    points.push({
      x: Math.random() * width,
      y: Math.random() * height,
      vx: (Math.random() - 0.5) * 0.34,
      vy: (Math.random() - 0.5) * 0.34,
      size: 1 + Math.random() * 1.8,
    });
  }
}

function draw() {
  ctx.clearRect(0, 0, width, height);

  for (const point of points) {
    point.x += point.vx;
    point.y += point.vy;

    if (point.x < 0 || point.x > width) point.vx *= -1;
    if (point.y < 0 || point.y > height) point.vy *= -1;
  }

  for (let i = 0; i < points.length; i += 1) {
    for (let j = i + 1; j < points.length; j += 1) {
      const a = points[i];
      const b = points[j];
      const dx = a.x - b.x;
      const dy = a.y - b.y;
      const distance = Math.sqrt(dx * dx + dy * dy);

      if (distance < 140) {
        const opacity = (1 - distance / 140) * 0.3;
        ctx.strokeStyle = `rgba(55, 215, 255, ${opacity})`;
        ctx.lineWidth = 1;
        ctx.beginPath();
        ctx.moveTo(a.x, a.y);
        ctx.lineTo(b.x, b.y);
        ctx.stroke();
      }
    }
  }

  for (const point of points) {
    ctx.fillStyle = "rgba(65, 230, 164, 0.72)";
    ctx.beginPath();
    ctx.arc(point.x, point.y, point.size, 0, Math.PI * 2);
    ctx.fill();
  }

  animationFrame = requestAnimationFrame(draw);
}

resize();
draw();
window.addEventListener("resize", () => {
  cancelAnimationFrame(animationFrame);
  resize();
  draw();
});
