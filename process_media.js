const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

const ffmpegPath = "C:\\Program Files\\AnyUkit\\ffmpeg.exe";
const srcDir = "D:\\antigravity\\1151123115國中數學領域數位科技推動自主學習探究\\Photos-1-001";
const photosOutDir = path.join(__dirname, 'assets', 'records', 'photos');
const videoOutDir = path.join(__dirname, 'assets', 'records', 'video');

if (!fs.existsSync(photosOutDir)) fs.mkdirSync(photosOutDir, { recursive: true });
if (!fs.existsSync(videoOutDir)) fs.mkdirSync(videoOutDir, { recursive: true });

console.log("Processing photos...");
const photoFiles = fs.readdirSync(srcDir).filter(f => f.toLowerCase().endsWith('.jpg') && f.startsWith('PXL_'));

photoFiles.sort();

console.log(`Found ${photoFiles.length} photos.`);
photoFiles.forEach((file, idx) => {
  const srcPath = path.join(srcDir, file);
  const outName = `record_photo_${String(idx + 1).padStart(2, '0')}.jpg`;
  const outPath = path.join(photosOutDir, outName);
  
  // Resize to max 1440 width/height, quality 85%
  const cmd = `"${ffmpegPath}" -y -i "${srcPath}" -vf "scale='min(1440,iw)':-2" -q:v 3 "${outPath}"`;
  console.log(`Compressing [${idx + 1}/${photoFiles.length}] ${file} -> ${outName}`);
  try {
    execSync(cmd, { stdio: 'pipe' });
    const inSize = (fs.statSync(srcPath).size / 1024).toFixed(1);
    const outSize = (fs.statSync(outPath).size / 1024).toFixed(1);
    console.log(`✓ ${outName}: ${inSize}KB -> ${outSize}KB`);
  } catch (err) {
    console.error(`Error processing ${file}:`, err.message);
  }
});

console.log("\nProcessing video IMG_8906.MOV...");
const srcVideo = path.join(srcDir, 'IMG_8906.MOV');
const outVideo = path.join(videoOutDir, 'teaching_demo_video.mp4');
const posterImg = path.join(videoOutDir, 'video_poster.jpg');

// Extract poster frame at 01:00
try {
  console.log("Extracting video poster image...");
  execSync(`"${ffmpegPath}" -y -ss 00:01:00 -i "${srcVideo}" -vframes 1 -q:v 2 "${posterImg}"`, { stdio: 'pipe' });
  console.log("✓ Poster created:", posterImg);
} catch (e) {
  console.error("Error generating poster:", e.message);
}

// Transcode video to 720p H.264 AAC web-optimized
console.log("Transcoding video to 720p H.264 (crf 28, faststart)...");
try {
  const videoCmd = `"${ffmpegPath}" -y -i "${srcVideo}" -vf "scale=-2:720" -c:v libx264 -crf 28 -preset fast -c:a aac -b:a 96k -movflags +faststart "${outVideo}"`;
  execSync(videoCmd, { stdio: 'inherit' });
  const inMB = (fs.statSync(srcVideo).size / (1024 * 1024)).toFixed(1);
  const outMB = (fs.statSync(outVideo).size / (1024 * 1024)).toFixed(1);
  console.log(`✓ Video successfully compressed: ${inMB}MB -> ${outMB}MB`);
} catch (e) {
  console.error("Error transcoding video:", e.message);
}

console.log("All media processing completed!");
