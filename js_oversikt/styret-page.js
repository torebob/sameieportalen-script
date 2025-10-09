<!-- public/js/ui/styret-page.html -->
<div class="space-y-6">
<h1 class="text-3xl font-bold text-gray-800">Styrets Oversikt</h1>
<p class="text-gray-600">Velkommen, styremedlem/styreleder. Her administrerer du seksjonseiere, avvik og årsmøter.</p>

<!-- Hurtiglenker for viktige funksjoner -->
<div class="grid grid-cols-1 md:grid-cols-3 gap-6">
<div class="bg-white p-6 rounded-xl shadow-lg border border-red-200">
<h3 class="text-xl font-semibold text-red-600">ÅPNE AVVIK</h3>
<p class="text-sm text-gray-500 mt-1">Status på alle innmeldte feil og mangler.</p>
<p class="text-4xl font-bold mt-4">3</p>
<a href="#avvik" class="text-red-500 hover:underline mt-2 inline-block">Gå til Avvik &rarr;</a>
</div>
<div class="bg-white p-6 rounded-xl shadow-lg border border-yellow-200">
<h3 class="text-xl font-semibold text-yellow-600">VENTENDE GODKJENNINGER</h3>
<p class="text-sm text-gray-500 mt-1">Fullmakter, refunderingskrav, leiesøknader.</p>
<p class="text-4xl font-bold mt-4">1</p>
<a href="#godkjenning" class="text-yellow-500 hover:underline mt-2 inline-block">Gå til Godkjenning &rarr;</a>
</div>
<div class="bg-white p-6 rounded-xl shadow-lg border border-blue-200">
<h3 class="text-xl font-semibold text-blue-600">NESTE ÅRSMØTE</h3>
<p class="text-sm text-gray-500 mt-1">Frist for saksinnmelding: 01. Mars 2026</p>
<p class="text-4xl font-bold mt-4">120</p>
<span class="text-blue-500">Dager igjen</span>
</div>
</div>

<!-- Avviksliste (Kun for demonstrasjon) -->
<div class="bg-white p-6 rounded-xl shadow-lg">
<h3 class="text-2xl font-semibold text-gray-800 mb-4">Siste Aktive Avvik</h3>
<ul class="divide-y divide-gray-200">
<li class="py-3 flex justify-between items-center">
<span class="font-medium">Heis ute av drift i bygg B</span>
<span class="text-sm bg-red-100 text-red-800 px-3 py-1 rounded-full">Kritisk</span>
</li>
<li class="py-3 flex justify-between items-center">
<span class="font-medium">Lyspære defekt i oppgang A</span>
<span class="text-sm bg-yellow-100 text-yellow-800 px-3 py-1 rounded-full">Middels</span>
</li>
</ul>
</div>
</div>

<script>
// Hvis du senere vil ha egen JavaScript-logikk for denne siden (f.eks. diagrammer):
// function initStyretPage() {
//    console.log('Styrets side lastet og initiert!');
// }
// window.onload = initStyretPage; // Merk: Dette må implementeres annerledes i en SPA.
</script>
