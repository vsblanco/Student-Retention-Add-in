import React, { useState, useEffect } from "react";

// updated user list with roles and optional photo (leave photo out or set to null to use initials)
const Users = [
	{ name: 'Victor Blanco', role: 'Student Services Coordinator', photo: 'https://wallpapers.com/images/featured/dolphin-w2b1iptrwaumv8de.jpg' },
	{ name: 'Angel Baez', role: 'Dean of Academic Affairs', photo: 'https://i.natgeofe.com/n/34b9d763-a5ef-434b-8d13-48f4919ca078/green-iguana_thumb_16x9.jpg?w=1200' },
	{ name: 'Darlen Gutierrez', role: 'Student Services Coordinator', photo: 'https://i.pinimg.com/736x/28/5e/5e/285e5e923048938dbe93d20d054c0c17.jpg' },
	{ name: 'Angel Coronel', role: 'Associate Dean of Academic Affairs', photo: 'https://play-lh.googleusercontent.com/RHT25lYamggGxosSgW5hUxphFluO4byH-1ZbOuW7nD1AK7QfuqS6ZK_fIoDdb8UIHw'},
	{ name: 'Kelvin Saliers', role: 'Full Time Instructor', photo: 'https://ftccollege.edu/wp-content/uploads/2023/08/mobile_programas_Construction_HVAC_R.jpg' },
	{ name: 'Yasser Rojas', role: 'Full Time Instructor', photo: 'https://assets.streamlinehq.com/image/private/w_512,h_512,ar_1/f_auto/v1/icons/freebies-freemojis/travel-places/travel-places/hospital-4e1hno0dnb9ndwe5yi08v.png?_a=DATAg1AAZAA0' },
];

export default function SSOtemp({ onSelect, defaultUser = null }) {
	// track the selected user
	const [selected, setSelected] = useState(defaultUser);

	useEffect(() => {
		if (selected && typeof onSelect === "function") {
			onSelect(selected);
		}
	}, [selected, onSelect]);

	// helper to get initials for avatar
	const initialsOf = (name) =>
		name
			.split(" ")
			.map((n) => n[0] || "")
			.slice(0, 2)
			.join("")
			.toUpperCase();

	return (
		<div className="relative p-0 flex-1 min-w-0 w-full shadow-xl rounded-lg bg-white overflow-hidden">
			{/* header */}
			<div className="px-6 py-5 bg-gradient-to-r from-sky-500 via-indigo-500 to-violet-600 shadow-sm">
				<div className="flex items-center gap-3">
					{/* icon */}
					<div className="flex items-center justify-center w-10 h-10 rounded-full bg-white/20 backdrop-blur-sm">
						<svg className="w-5 h-5 text-white" viewBox="0 0 24 24" fill="none" aria-hidden>
							<path d="M12 12c2.761 0 5-2.239 5-5s-2.239-5-5-5-5 2.239-5 5 2.239 5 5 5z" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
							<path d="M21 21c0-2.761-4.03-5-9-5s-9 2.239-9 5" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
						</svg>
					</div>

					{/* title + subtitle */}
					<div>
						<h3 className="text-white text-lg font-semibold leading-tight">Select User</h3>
						<p className="text-white/85 text-xs mt-0.5">Choose an account to comment as</p>
					</div>
				</div>
			</div>

			<div className="p-5">
				<div id="user-buttons-container" className="space-y-3" role="list">
					{Users.map((user, idx) => (
	<button
		key={user.name}
		type="button"
		role="listitem"
		aria-pressed={selected === user.name}
		className={
			// make the button a group and positioned so overlay can sit on top
			`relative group overflow-hidden flex items-center gap-3 w-full text-left px-4 py-3 rounded-md transition-shadow duration-200 ` +
			(selected === user.name
				? "bg-blue-50 shadow-sm ring-1 ring-blue-200"
				: "bg-white hover:shadow-lg")
		}
		style={{ ['--delay']: `${idx * 80}ms` }}
		onClick={() => setSelected(user.name)}
		>
		{/* avatar */}
		<div
			className={`flex-none w-10 h-10 rounded-full flex items-center justify-center font-semibold text-sm text-white overflow-hidden ${user.photo && selected === user.name ? 'ring-2 ring-offset-1 ring-blue-400' : ''}`}
			style={{
				// only apply background when there's no photo
				background: !user.photo ? (selected === user.name ? "linear-gradient(135deg,#2563eb,#7c3aed)" : "#e5e7eb") : undefined
			}}
		>
			{user.photo ? (
				<img src={user.photo} alt={user.name} className="w-full h-full object-cover" />
			) : (
				<span style={{ color: selected === user.name ? "#fff" : "#374151" }}>{initialsOf(user.name)}</span>
			)}
		</div>

		{/* name/details */}
		<div className="flex-1">
			<div className="text-sm font-medium text-gray-800">{user.name}</div>
			<div className="text-xs text-gray-500">{user.role}</div>
		</div>

		{/* selected check */}
		<div className="flex-none w-6 h-6">
			{selected === user.name ? (
				<svg width="20" height="20" viewBox="0 0 24 24" fill="none" aria-hidden>
					<path d="M20 6L9 17l-5-5" stroke="#2563eb" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round" />
				</svg>
			) : null}
		</div>

		{/* slight dark overlay on hover (non-interactive) */}
		<span aria-hidden className="absolute inset-0 rounded-md bg-black opacity-0 transition-opacity duration-200 group-hover:opacity-10 pointer-events-none"></span>
	</button>
))}
				</div>
			</div>

			{/* animation styles (staggered fade-in) */}
			<style>{`
				#user-buttons-container > button {
					opacity: 0;
					transform: translateY(8px);
					animation: fadeInUp 420ms cubic-bezier(.2,.9,.3,1) both;
					animation-delay: var(--delay, 0ms);
				}
				@keyframes fadeInUp {
					from { opacity: 0; transform: translateY(8px); }
					to   { opacity: 1; transform: translateY(0); }
				}
				/* small responsive tweak */
				@media (max-width: 420px) {
					:root { --container-w: 100%; }
				}
			`}</style>
		</div>
	);
}
